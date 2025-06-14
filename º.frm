VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIWork 
   BackColor       =   &H008080FF&
   Caption         =   "AuditCall"
   ClientHeight    =   8304
   ClientLeft      =   108
   ClientTop       =   -3216
   ClientWidth     =   16176
   Icon            =   "º.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PctMdi 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12000
      Left            =   0
      ScaleHeight     =   11952
      ScaleWidth      =   16128
      TabIndex        =   0
      Top             =   0
      Width           =   16176
      Begin VB.Image Image1 
         Height          =   5832
         Left            =   3756
         Picture         =   "º.frx":000C
         Top             =   1512
         Width           =   7668
      End
   End
   Begin Crystal.CrystalReport CrpRep 
      Left            =   0
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu MnuMaster 
      Caption         =   "&1 Master"
      Begin VB.Menu MnuGenMst 
         Caption         =   "&1 General Master"
         Begin VB.Menu MnuComp 
            Caption         =   "&0 About Us"
         End
         Begin VB.Menu MnuGroup 
            Caption         =   "&1 Entity Nature"
         End
         Begin VB.Menu MnuClientList 
            Caption         =   "&2 Client List"
         End
         Begin VB.Menu MnuHeadList 
            Caption         =   "&3 Head List"
         End
         Begin VB.Menu MnuLedMst 
            Caption         =   "&4 Ledger"
         End
         Begin VB.Menu MnuLedDtl 
            Caption         =   "&5 Ledger Detail"
         End
         Begin VB.Menu MnuMemberMst 
            Caption         =   "&6 Membership No."
         End
      End
      Begin VB.Menu MnuBSheetList 
         Caption         =   "&2 Balance sheet (Schedule VIII) List"
      End
      Begin VB.Menu MnuIncExpList 
         Caption         =   "&3 Income Expenditure A/c. (Schedule IX) List"
      End
      Begin VB.Menu MnuSchMemoMst 
         Caption         =   "&4 School Memo"
      End
      Begin VB.Menu MnuSchedule9 
         Caption         =   "&5 Schedule IX-C"
      End
   End
   Begin VB.Menu MnuTran 
      Caption         =   "&2 Transaction"
      Begin VB.Menu MnuRecPay 
         Caption         =   "&1 Receipts And Payments"
      End
      Begin VB.Menu MnuIncExpTrn 
         Caption         =   "&2 Income Expenditure Account"
      End
      Begin VB.Menu MnuBalanceSheetTrn 
         Caption         =   "&3 Balance Sheet"
      End
      Begin VB.Menu MnuJv 
         Caption         =   "&4 Journal Entry"
      End
      Begin VB.Menu MnuSchMemoTrn 
         Caption         =   "&5 School Memo"
      End
      Begin VB.Menu MnuAnlRep 
         Caption         =   "&6 Annual Report"
      End
      Begin VB.Menu MnuConstrn 
         Caption         =   "&7 Consolidated"
         Begin VB.Menu MnuIncExpConTrn 
            Caption         =   "&1 Income Expenditure Account"
         End
         Begin VB.Menu MnuBSConTrn 
            Caption         =   "&2 Balance Sheet"
         End
         Begin VB.Menu MnuSch9CEntry 
            Caption         =   "&3 Schedule 9C"
         End
      End
   End
   Begin VB.Menu MnuReport 
      Caption         =   "&3 Report"
      Begin VB.Menu MnuAuditReport 
         Caption         =   "&1 Audit Report"
      End
      Begin VB.Menu MnuBSReport 
         Caption         =   "&2 Balance Sheet"
         Begin VB.Menu MnuBsConsoliRep 
            Caption         =   "&1 Consolidated"
         End
         Begin VB.Menu MnuBsIndRep 
            Caption         =   "&2 Independent"
         End
         Begin VB.Menu MnuBsNUdinBsRep 
            Caption         =   "&3 Non-UDIN"
         End
      End
      Begin VB.Menu MnuIncExpReport 
         Caption         =   "&3 Income Expenditure Statement"
         Begin VB.Menu MnuIncExpConReport 
            Caption         =   "&1 Consolidated"
         End
         Begin VB.Menu MnuIncExpIndReport 
            Caption         =   "&2 Independent"
         End
         Begin VB.Menu MnuIncExpNUReport 
            Caption         =   "&3 Non-UDIN"
         End
      End
      Begin VB.Menu MnuRecPayReport 
         Caption         =   "&4 Receipts Payments Account "
         Begin VB.Menu MnuRecPayIndReport 
            Caption         =   "&1 Independent"
         End
         Begin VB.Menu MnuRecPayIndvUReport 
            Caption         =   "&2 Individual UDIN"
         End
         Begin VB.Menu MnuRecPayIndvNReport 
            Caption         =   "&3 Non-UDIN"
         End
      End
      Begin VB.Menu MnuSchoolMemoReport 
         Caption         =   "&5 School Memo"
      End
      Begin VB.Menu MnuFCStateReport 
         Caption         =   "&6 FC Certificate"
      End
      Begin VB.Menu MnuLibraryFormReport 
         Caption         =   "&7 Library Form"
      End
   End
   Begin VB.Menu MnuWinU 
      Caption         =   "&4 Window"
      WindowList      =   -1  'True
      Begin VB.Menu MnuTileH 
         Caption         =   "&1 Tile Horizontally"
      End
      Begin VB.Menu MnuTileV 
         Caption         =   "&2 Tile Vertically"
      End
      Begin VB.Menu MnuCascade 
         Caption         =   "&3 Cascade"
      End
   End
   Begin VB.Menu MnuUtility 
      Caption         =   "&5 Utility"
      Begin VB.Menu MnuUtiUser 
         Caption         =   "&1 User Admin"
      End
      Begin VB.Menu MnuOpenBal 
         Caption         =   "&2 Opening Balance"
      End
   End
   Begin VB.Menu MnuExit 
      Caption         =   "&6 Exit"
   End
End
Attribute VB_Name = "MDIWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    Me.Caption = mCmpName
    SetMenu
End Sub

Private Sub MnuAnlRep_Click()
    PctMdi.Visible = False
    FrmAnnualRMst.Show
End Sub

Private Sub MnuAuditReport_Click()
    PctMdi.Visible = False
    FrmAuditRepPrint.Show
End Sub

Private Sub MnuBalanceSheetTrn_Click()
    PctMdi.Visible = False
    FrmBS.Show
End Sub

Private Sub MnuBsConsoliRep_Click()
    PctMdi.Visible = False
    FrmBSConRep.Show
End Sub

Private Sub MnuBSConTrn_Click()
    PctMdi.Visible = False
    FrmBSConEntry.Show
End Sub

Private Sub MnuBSheetList_Click()
    PctMdi.Visible = False
    FrmBSHeadMst.Show
End Sub

Private Sub MnuBsIndRep_Click()
    PctMdi.Visible = False
    FrmBSIndRep.Show
End Sub

Private Sub MnuBsNUdinBsRep_Click()
    PctMdi.Visible = False
    FrmBSNUDINRep.Show
End Sub

Private Sub MnuCascade_Click()
    Me.Arrange 0
End Sub

Private Sub MnuClientList_Click()
    PctMdi.Visible = False
    FrmClientMst.Show
End Sub

Private Sub MnuComp_Click()
    PctMdi.Visible = False
    FrmCompany.Show
End Sub
Private Sub MnuExit_Click()
    PctMdi.Visible = False
    FrmShut.Show
End Sub

Private Sub MnuFCStateReport_Click()
    PctMdi.Visible = False
    FrmFCReport.Show
End Sub

Private Sub MnuGroup_Click()
    PctMdi.Visible = False
    FrmEntityNature.Show
End Sub

Private Sub MnuHeadList_Click()
    PctMdi.Visible = False
    FrmHeadMst.Show
End Sub

Private Sub MnuIncExpConReport_Click()
    PctMdi.Visible = False
    FrmPlAcConRep.Show
End Sub

Private Sub MnuIncExpConTrn_Click()
    PctMdi.Visible = False
    FrmPlAcConEntry.Show
End Sub

Private Sub MnuIncExpIndReport_Click()
    PctMdi.Visible = False
    FrmPlAcIndRep.Show
End Sub

Private Sub MnuIncExpList_Click()
    PctMdi.Visible = False
    FrmPLHeadMst.Show
End Sub

Private Sub MnuIncExpNUReport_Click()
    PctMdi.Visible = False
    FrmPlAcNUDINRep.Show
End Sub

Private Sub MnuIncExpTrn_Click()
    PctMdi.Visible = False
    FrmPlAc.Show
End Sub

Private Sub MnuJv_Click()
    PctMdi.Visible = False
    FrmJV.Show
End Sub

Private Sub MnuLedDtl_Click()
    PctMdi.Visible = False
    FrmSubLedMst.Show
End Sub

Private Sub MnuLedMst_Click()
    PctMdi.Visible = False
    FrmLedMst.Show
End Sub

Private Sub MnuMemberMst_Click()
    PctMdi.Visible = False
    FrmMemberMst.Show
End Sub

Private Sub MnuOpenBal_Click()
    PctMdi.Visible = False
    FrmUOpenBal.Show
End Sub

Private Sub MnuRecPay_Click()
    PctMdi.Visible = False
    FrmRecPay.Show
End Sub

Private Sub MnuRecPayIndReport_Click()
    PctMdi.Visible = False
    FrmRecPayIDRep.Show
End Sub

Private Sub MnuRecPayIndvNReport_Click()
    PctMdi.Visible = False
    FrmRecPayNUDINRep.Show
End Sub

Private Sub MnuSch9CEntry_Click()
    PctMdi.Visible = False
    FrmSchedule9CEntry.Show
End Sub

Private Sub MnuSchedule9_Click()
    PctMdi.Visible = False
    FrmSed9Mst.Show
End Sub

Private Sub MnuSchMemoMst_Click()
    PctMdi.Visible = False
    FrmSMHeadMst.Show
End Sub

Private Sub MnuSchMemoTrn_Click()
    PctMdi.Visible = False
    FrmSchoolMemo.Show
End Sub

Private Sub MnuSchoolMemoReport_Click()
    PctMdi.Visible = False
    FrmSchoolMemoRep.Show
End Sub

Private Sub MnuTileH_Click()
    Me.Arrange 0
    Me.Arrange 2
End Sub
Private Sub MnuTileV_Click()
    Me.Arrange 0
    Me.Arrange 1
End Sub

Private Sub MnuUtiUser_Click()
    PctMdi.Visible = False
    FrmUser.Show
End Sub
Private Sub SetMenu()
If mUType <> "Admin" Then
    MnuComp.Enabled = False
    MnuGroup.Enabled = False
    MnuBSheetList.Enabled = False
    MnuHeadList.Enabled = False
    MnuIncExpList.Enabled = False
    MnuLedDtl.Enabled = False
    MnuLedMst.Enabled = False
    MnuMemberMst.Enabled = False
    MnuSchedule9.Enabled = False
    MnuSchMemoMst.Enabled = False
    MnuUtiUser.Enabled = False
End If
End Sub
