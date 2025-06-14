VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmSed9Mst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule 9 C"
   ClientHeight    =   8865
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   15705
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FS9Mst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   15705
   Begin VB.Frame FraMst 
      Height          =   7695
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   15492
      Begin VB.ComboBox ComNature 
         Height          =   324
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2652
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   6972
         Left            =   120
         TabIndex        =   2
         Top             =   636
         Width           =   15252
         _cx             =   26903
         _cy             =   12298
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
         FormatString    =   $"FS9Mst.frx":0442
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
         ShowComboButton =   -1  'True
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
         Caption         =   "Firm Nature :"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1248
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   7572
         Left            =   0
         Top             =   120
         Width           =   15492
      End
   End
   Begin VB.Frame FraTool 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   6135
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
         Left            =   5160
         Picture         =   "FS9Mst.frx":048B
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "FS9Mst.frx":08CD
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "FS9Mst.frx":0D0F
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "FS9Mst.frx":1379
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "FS9Mst.frx":17BB
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "FS9Mst.frx":1BFD
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Add"
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape ShpMain 
         Height          =   975
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   6135
      End
   End
End
Attribute VB_Name = "FrmSed9Mst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mActivity As String
Private Sub ComNature_Validate(Cancel As Boolean)
    If mActivity <> "" Then ComNature.Locked = True Else ComNature.Locked = False
    SetData
    Display
End Sub
Private Sub Form_Load()
    Me.Left = 50
    Me.Top = 50
    SetCombo
    SetData
    Display
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select * From GrpMst Where GCode=1", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
ComNature.Clear
Do While RsQry.EOF = False
    ComNature.AddItem RsQry.Fields("GName")
    ComNature.ItemData(ComNature.NewIndex) = RsQry.Fields("GCode")
    RsQry.MoveNext
Loop
ComNature.ListIndex = 0
End Function
Private Sub ClearText()
    ComNature.ListIndex = 0
End Sub
Private Function SaveData()
Dim i As Double
On Error GoTo XErr
DbWorkAuto.BeginTrans
DbWorkAuto.Execute "Delete From S9HeadTrn Where GCode=" & ComNature.ItemData(ComNature.ListIndex)
If mActivity <> "Delete" Then
With VsfHelp
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 0)) <> 0 And Val(.TextMatrix(i, 6)) <> 0 Then
            DbWorkAuto.Execute "Insert InTo S9HeadTrn (GroupCode,SrN,GCode) Values (" & Val(.TextMatrix(i, 6)) & "," & _
            Val(.TextMatrix(i, 0)) & "," & ComNature.ItemData(ComNature.ListIndex) & ")"
        End If
    Next
End With
End If
DbWorkAuto.CommitTrans
Exit Function
XErr:
MsgBox Err.Description
DbWorkAuto.RollbackTrans
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
Private Sub VsfHelp_EnterCell()
    If mActivity = "Add" Or mActivity = "Edit" Then
        If VsfHelp.Col = 0 Then
            VsfHelp.Editable = flexEDKbd
            SendKeys "{F2}"
        Else
            VsfHelp.Editable = flexEDNone
        End If
    End If
End Sub
Private Sub VsfHelp_RowColChange()
    If mActivity = "Add" Or mActivity = "Edit" Then
        If VsfHelp.Col = 0 Then
            VsfHelp.Editable = flexEDKbd
            SendKeys "{F2}"
        Else
            VsfHelp.Editable = flexEDNone
        End If
    End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Add"
        mActivity = "Add"
        SetTool False
        ClearText
        VsfHelp.Enabled = True
        ComNature.SetFocus
    Case "Edit"
        If VsfHelp.Rows > 1 Then
            VsfHelp.Enabled = True
            mActivity = "Edit"
            SetTool False
            ComNature.SetFocus
        Else
            MsgBox "Sorry !! Not Allowded..", vbInformation, "Black Data Error"
        End If
    Case "Delete"
        If VsfHelp.Rows > 1 Then
            VsfHelp.Enabled = True
            mActivity = "Delete"
            SetTool False
            ComNature.SetFocus
        Else
            MsgBox "Sorry !! Not Allowded..", vbInformation, "Black Data Error"
        End If
    Case "Save"
        If ComNature.Text = "" Then
            MsgBox "Sorry !! Not Allowded..", vbInformation, "Black Data Error"
            ComNature.SetFocus
        Else
            SaveData
            mActivity = ""
            ClearText
            VsfHelp.Enabled = False
            SetCombo
            SetTool True
            SetData
            Display
            ComNature.Locked = False
            ComNature.SetFocus
        End If
    Case "Cancel"
        mActivity = ""
        VsfHelp.Enabled = False
        SetTool True
        ClearText
        SetCombo
        SetTool True
        SetData
        ComNature.Locked = False
        Display
        ComNature.SetFocus
    Case "Exit"
        Unload Me
End Select
End Sub

Private Function SetTool(ByVal mVal As Boolean)
TlbSav(0).Enabled = mVal
TlbSav(1).Enabled = mVal
TlbSav(2).Enabled = mVal
TlbSav(3).Enabled = Not mVal
If mUType <> "A" Then
    TlbSav(0).Enabled = False
    TlbSav(1).Enabled = False
    TlbSav(2).Enabled = False
End If
End Function

Private Sub SetData()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select '' As SrN,EntMst.EName,LedMst.LName,HedMst.HName,IIF(HedMst.HType=1,'Balance Sheet','Income And Expenditure')" & _
" As HeadType,HedMst.HSide,EntMst.ECode From EntMst,LedMst,HedMst Where EntMst.LCode=LedMst.LCode And LedMst.HCode=HedMst.HCode Order By EntMst.EName,LedMst.LName", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfHelp.DataSource = RsQry
With VsfHelp
    .FontSize = 11
    .TextMatrix(0, 0) = "SR"
    .ColWidth(0) = 450
    .TextMatrix(0, 1) = "Sub Ledger"
    .ColWidth(1) = 3800
    .TextMatrix(0, 2) = "Ledger"
    .ColWidth(2) = 4800
    .TextMatrix(0, 3) = "Under Head"
    .ColWidth(3) = 3200
    .TextMatrix(0, 4) = "Main Head"
    .ColWidth(4) = 2200
    .TextMatrix(0, 5) = "Side"
    .ColWidth(5) = 400
    .ColWidth(6) = 0    'GROUPCODE
    .Refresh
End With
End Sub

Private Sub Display()
Dim RsQry As New ADODB.Recordset
Dim i As Double
RsQry.Open "Select * From SeqMst Where GCode=" & ComNature.ItemData(ComNature.ListIndex) & " Order By GSrN", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
With VsfHelp
    Do While RsQry.EOF = False
        .Row = 1
        i = .FindRow(RsQry.Fields("HCode"), 1, 6)
        If i > 0 Then .TextMatrix(i, 0) = RsQry.Fields("GSrN")
        RsQry.MoveNext
    Loop
End With
End Sub
