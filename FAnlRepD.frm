VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmUnitRptDtl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Report Detail"
   ClientHeight    =   8790
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   19590
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FAnlRepD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   19590
   Begin VB.Frame FraAcHelp 
      Height          =   6612
      Left            =   28000
      TabIndex        =   21
      Top             =   1680
      Width           =   12132
      Begin VB.CommandButton CmdAClose 
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exit"
         Top             =   6120
         Width           =   972
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfAcHelp 
         Height          =   5772
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   11892
         _cx             =   20976
         _cy             =   10181
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   7.5
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
         BackColorBkg    =   16777152
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
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FAnlRepD.frx":0442
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
         Height          =   6492
         Index           =   1
         Left            =   0
         Top             =   120
         Width           =   12132
      End
   End
   Begin VB.Frame FraMst 
      Height          =   7692
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   19335
      Begin VB.TextBox TxtOMetter 
         Height          =   1950
         Left            =   1560
         TabIndex        =   8
         Top             =   5520
         Width           =   8010
      End
      Begin VB.TextBox TxtEmphasis 
         Height          =   1830
         Left            =   1560
         TabIndex        =   7
         Top             =   3600
         Width           =   8010
      End
      Begin VB.ComboBox ComQualify 
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   1440
      End
      Begin VB.ComboBox ComFAsset 
         Height          =   360
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2520
      End
      Begin VB.ComboBox ComAcMethod 
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2400
      End
      Begin VB.TextBox TxtQualify 
         Height          =   1830
         Left            =   1560
         TabIndex        =   6
         Top             =   1680
         Width           =   8010
      End
      Begin VB.TextBox TxtName 
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   5085
      End
      Begin VB.TextBox TxtFileNo 
         Height          =   384
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1410
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   7335
         Left            =   9600
         TabIndex        =   0
         Top             =   270
         Width           =   9615
         _cx             =   16960
         _cy             =   12938
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
         BackColor       =   16777215
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
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FAnlRepD.frx":048B
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
         Caption         =   "Other Matters :"
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
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   5520
         Width           =   1410
      End
      Begin VB.Label LblCompany 
         Caption         =   "Emphasis of Matter:"
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
         Height          =   600
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label LblCompany 
         Caption         =   "Qualified Report:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1680
      End
      Begin VB.Label LblCompany 
         Caption         =   "Fixed Assets Inventory :"
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
         Left            =   4680
         TabIndex        =   25
         Top             =   720
         Width           =   2325
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   7575
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   19335
      End
      Begin VB.Label LblCompany 
         Caption         =   "Qualification:"
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
         TabIndex        =   19
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label LblCompany 
         Caption         =   "Accounting Method :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label LblCompany 
         Caption         =   "File No.:"
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
         Height          =   228
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   912
      End
   End
   Begin VB.Frame FraTool 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   7092
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
         Height          =   735
         Index           =   6
         Left            =   5160
         Picture         =   "FAnlRepD.frx":04D4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   855
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
         Height          =   735
         Index           =   5
         Left            =   6120
         Picture         =   "FAnlRepD.frx":0B3E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exit"
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
         Height          =   735
         Index           =   4
         Left            =   4080
         Picture         =   "FAnlRepD.frx":0F80
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancel"
         Top             =   240
         Width           =   975
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
         Height          =   735
         Index           =   3
         Left            =   3030
         Picture         =   "FAnlRepD.frx":13C2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   975
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
         Height          =   735
         Index           =   2
         Left            =   1890
         Picture         =   "FAnlRepD.frx":1A2C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Delete"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   1005
         Picture         =   "FAnlRepD.frx":1E6E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Edit"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   120
         Picture         =   "FAnlRepD.frx":22B0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Add"
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape ShpMain 
         Height          =   972
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   7092
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
   Begin VB.Label LblCompany 
      Caption         =   "Press F4 Key For File No. Help"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   432
      Index           =   26
      Left            =   10320
      TabIndex        =   20
      Top             =   480
      Width           =   4956
   End
End
Attribute VB_Name = "FrmUnitRptDtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mActivity As String
Dim mAcCode As Double
Private Sub CmdAClose_Click()
    FraAcHelp.Left = 28000
    TxtFileNo.SetFocus
End Sub
Private Sub Form_Load()
    Me.Left = 50
    Me.Top = 50
    SetCombo
    If VsfHelp.Rows > 1 Then Display
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.FileNo,AcMst.AcName,ArpDtl.* From AcMst,ArpDtl Where AcMst.AcCode=ArpDtl.AcCode And AcMst.Active=-1 And AcMst.PACode<>0 Order By FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfHelp.DataSource = RsQry
With VsfHelp
    .TextMatrix(0, 0) = "FILE NO"
    .ColWidth(0) = 1300
    .TextMatrix(0, 1) = "CLIENT NAME"
    .ColWidth(1) = 6400
    .ColWidth(2) = 0    'AcCode
    .ColWidth(3) = 500    'AcBasce
    .ColWidth(4) = 500    'FixAsset
    .ColWidth(5) = 500    'Qua
    .ColWidth(6) = 0    'Qua Text
    .ColWidth(7) = 0    'Emphasis
    .ColWidth(8) = 0    'OMatter
    .Refresh
End With
Set RsQry = Nothing
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode,'' As RName From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 And " & _
"AcMst.AcType=GrpMst.GCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.City,AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,''),GrpMst.GName,AcMst.AcCode," & _
"'' From AcMst,GrpMst,AcMst As AcMst1 Where AcMst.Active=-1 And AcMst.PaCode=AcMst1.AcCode And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfAcHelp.DataSource = RsQry
With VsfAcHelp
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
End With
ComAcMethod.Clear
ComAcMethod.AddItem "Cash Basis"
ComAcMethod.ItemData(ComAcMethod.NewIndex) = 1
ComAcMethod.AddItem "Mercantile Basis"
ComAcMethod.ItemData(ComAcMethod.NewIndex) = 2
ComAcMethod.ListIndex = 0
ComFAsset.Clear
ComFAsset.AddItem "Not Prepared"
ComFAsset.ItemData(ComFAsset.NewIndex) = 1
ComFAsset.AddItem "Prepared"
ComFAsset.ItemData(ComFAsset.NewIndex) = 2
ComFAsset.ListIndex = 0
ComQualify.Clear
ComQualify.AddItem "No"
ComQualify.ItemData(ComQualify.NewIndex) = 1
ComQualify.AddItem "Yes"
ComQualify.ItemData(ComQualify.NewIndex) = 2
ComQualify.ListIndex = 0
End Function
Private Sub ClearText()
    TxtEmphasis.Text = ""
    TxtFileNo.Text = ""
    TxtName.Text = ""
    TxtOMetter.Text = ""
    TxtQualify.Text = ""
    ComAcMethod.ListIndex = 0
    ComFAsset.ListIndex = 0
    ComQualify.ListIndex = 0
End Sub
Private Function SaveData()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From ArpDtl Where AcCode=" & mAcCode
If mActivity <> "Delete" Then
    DbDataDB.Execute "Insert InTo ArpDtl (AcCode,AcBase,FAInvt,Qualif,QulRpt,EOMRpt,AOMRpt) Values (" & mAcCode & "," & ComAcMethod.ItemData(ComAcMethod.ListIndex) & "," & _
    IIf(ComFAsset.ItemData(ComFAsset.ListIndex) = 1, 0, -1) & "," & IIf(ComQualify.ItemData(ComQualify.ListIndex) = 1, 0, -1) & ",'" & TxtQualify.Text & "','" & _
    TxtEmphasis.Text & "','" & TxtOMetter.Text & "')"
End If
DbDataDB.CommitTrans
Exit Function
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mActivity = "Add" Or mActivity = "Edit" Or mActivity = "Delete" Then
    If MsgBox("The Activity " + mActivity + " Is Not Saved " + vbCrLf + "Do You Want To Exit ? ", vbCritical + vbYesNo, "Exit Error") = vbYes Then
        mActivity = ""
        MDIWork.PctMdi.Visible = True
        Unload Me
    Else
        Cancel = 1
    End If
Else
    mActivity = ""
    MDIWork.PctMdi.Visible = True
    Unload Me
End If
End Sub
Private Sub TxtFileNo_GotFocus()
    If mActivity <> "" And TxtFileNo.Text = "" Then
        If VsfAcHelp.Rows > 1 Then
            FraAcHelp.Left = 1080
            VsfAcHelp.SetFocus
        End If
    End If
End Sub

Private Sub TxtFileNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    If mActivity <> "" Then
        If VsfAcHelp.Rows > 1 Then
            FraAcHelp.Left = 1080
            VsfAcHelp.SetFocus
        End If
    End If
End If
End Sub

Private Sub TxtQualify_GotFocus()
    If ComQualify.Text = "Yes" Then TxtQualify.Locked = False Else TxtQualify.Locked = True
End Sub
Private Sub VsfHelp_RowColChange()
    Display
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Add"
        mActivity = "Add"
        SetTool False
        ClearText
        VsfHelp.Enabled = False
        TxtFileNo.SetFocus
    Case "Edit"
        If VsfHelp.Rows > 1 Then
            VsfHelp.Enabled = True
            mActivity = "Edit"
            SetTool False
            TxtFileNo.SetFocus
        Else
            MsgBox "Sorry! Not Allowed.", vbInformation, "Black Data Error"
        End If
    Case "Delete"
        If VsfHelp.Rows > 1 Then
            VsfHelp.Enabled = False
            mActivity = "Delete"
            SetTool False
            TxtFileNo.SetFocus
        Else
            MsgBox "Sorry! Not Allowed.", vbInformation, "Black Data Error"
        End If
    Case "Save"
        If TxtFileNo.Text = "" Then
            MsgBox "Sorry! Please Select File No.", vbInformation, "Black Data Error"
            TxtFileNo.SetFocus
        ElseIf DupliRec = True Then
            MsgBox "Sorry! Duplicate Record Found.", vbInformation, "Black Data Error"
            TxtFileNo.SetFocus
        Else
            SaveData
            DbDataDB.BeginTrans
                If mActivity = "Add" Then
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'UNIT_RPT','ADD_NEW','" & Date & "','" & Time & "')"
                ElseIf mActivity = "Delete" Then
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'UNIT_RPT','DELETE','" & Date & "','" & Time & "')"
                ElseIf mActivity = "Edit" Then
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'UNIT_RPT','UPDATE','" & Date & "','" & Time & "')"
                End If
            DbDataDB.CommitTrans
            mActivity = ""
            ClearText
            VsfHelp.Enabled = True
            SetCombo
            SetTool True
            If VsfHelp.Rows > 1 Then Display
            VsfHelp.SetFocus
        End If
    Case "Print"
        If mAcCode = 0 Then
            MsgBox "Please select client.", vbCritical, "Alert"
            Exit Sub
        End If
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'UNIT_RPT','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
    Case "Cancel"
        mActivity = ""
        VsfHelp.Enabled = True
        SetTool True
        ClearText
        SetCombo
        SetTool True
        If VsfHelp.Rows > 1 Then Display
        VsfHelp.SetFocus
    Case "Exit"
        Unload Me
End Select
End Sub

Private Function SetTool(ByVal mVal As Boolean)
TlbSav(0).Enabled = mVal
TlbSav(1).Enabled = mVal
TlbSav(2).Enabled = mVal
TlbSav(3).Enabled = Not mVal
End Function

Private Sub Display()
With VsfHelp
    TxtFileNo.Text = .TextMatrix(.Row, 0)
    TxtName.Text = .TextMatrix(.Row, 1)
    mAcCode = Val(.TextMatrix(.Row, 2))
    ComAcMethod.ListIndex = IIf(Val(.TextMatrix(.Row, 3)) = 2, 1, 0)
    ComFAsset.ListIndex = IIf(Val(.TextMatrix(.Row, 4)) = -1, 1, 0)
    ComQualify.ListIndex = IIf(Val(.TextMatrix(.Row, 5)) = -1, 1, 0)
    TxtQualify.Text = .TextMatrix(.Row, 6)
    TxtEmphasis.Text = .TextMatrix(.Row, 7)
    TxtOMetter.Text = .TextMatrix(.Row, 8)
    Dim i As Double
    i = VsfAcHelp.FindRow(TxtFileNo.Text, 1, 1)
    If i > -1 Then VsfAcHelp.Row = i
End With
End Sub
Private Sub VsfAcHelp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TxtName.Text = VsfAcHelp.TextMatrix(VsfAcHelp.Row, 0)
        TxtFileNo.Text = VsfAcHelp.TextMatrix(VsfAcHelp.Row, 1)
        mAcCode = Val(VsfAcHelp.TextMatrix(VsfAcHelp.Row, 5))
        FraAcHelp.Left = 28000
        TxtFileNo.SetFocus
    End If
End Sub

Private Function DupliRec() As Boolean
Dim RsQ As New ADODB.Recordset
If mAcCode <> Val(VsfHelp.TextMatrix(VsfHelp.Row, 2)) Then
    RsQ.Open "Select * From ArpDtl Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.EOF = False Then DupliRec = True Else DupliRec = False
Else
    DupliRec = False
End If
End Function

Private Sub PrintRec()
Dim RsRep As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim i As Double
Set RsRep = Nothing
RsRep.Open "Select RepDtl.*,AdtMst.AdtName From RepDtl,AdtMst Where AcCode=" & mAcCode & " And RepDtl.AdtNo=AdtMst.AdtNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsRep.EOF = False Then
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        MsgBox "UDIN not issued. Report cannot be printed.", vbInformation, "Information"
        Exit Sub
    End If
Else
    MsgBox "Signing Information not available. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
RsQ.Open "Select * From AcMst Where AcCode=(Select SaCode From QGroup Where AcCode=" & mAcCode & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\AdtRpt.Rpt"
    .SelectionFormula = "{TranBr.RepName}='VPDUW\'"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(2) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(3) = "mTitle1='Opinion on Financial Statements for the year ended on " & RsRep.Fields("RtDt") & "'"
    .Formulas(4) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(5) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
    .Formulas(6) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(7) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(8) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(9) = "mTUnit='" & VsfAcHelp.TextMatrix(VsfAcHelp.Row, 0) & IIf(IsNull(VsfAcHelp.TextMatrix(VsfAcHelp.Row, 2)) = False, ", " & VsfAcHelp.TextMatrix(VsfAcHelp.Row, 2), "") & "'"
    .Formulas(10) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(11) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(12) = "mTName='" & RsQ.Fields("AcName") & ", " & RsQ.Fields("City") & "'"
    .Formulas(13) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(14) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(15) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(16) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(17) = "mRTDate='" & RsRep.Fields("RtDt") & "'"
    If ComQualify.Text = "Yes" Then .Formulas(18) = "mQualif='Qualified Opinion'" Else .Formulas(18) = "mQualif='Opinion'"
    If ComAcMethod.Text = "Cash Basis" Then .Formulas(19) = "mAcBase='cash basis'" Else .Formulas(19) = "mAcBase='mercantile basis'"
    If ComFAsset.Text = "Prepared" Then .Formulas(20) = "mFAInvt='An'" Else .Formulas(20) = "mFAInvt='No'"
    .Formulas(21) = "mCType='U'"
    .Action = 1
    For i = 0 To 21
        .Formulas(i) = ""
    Next
    .SelectionFormula = ""
End With
End Sub
