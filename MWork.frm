VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIWork 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10635
   ClientLeft      =   105
   ClientTop       =   -5925
   ClientWidth     =   20250
   Icon            =   "MWork.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PctMdi 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12000
      Left            =   0
      Picture         =   "MWork.frx":2872A
      ScaleHeight     =   11940
      ScaleWidth      =   20190
      TabIndex        =   0
      Top             =   0
      Width           =   20250
   End
   Begin Crystal.CrystalReport CrpRep 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu MnuMaster 
      Caption         =   "&1 Utility"
      Begin VB.Menu MnuClient 
         Caption         =   "&1 Client List"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOpenBal 
         Caption         =   "&2 Opening Balance"
      End
      Begin VB.Menu MnuSoftUp 
         Caption         =   "&3 Software Update"
      End
   End
   Begin VB.Menu MnuTran 
      Caption         =   "&2 Data Entry"
      Begin VB.Menu MnuRecPay 
         Caption         =   "&1 Receipts And Payments"
      End
      Begin VB.Menu MnuJv 
         Caption         =   "&2 Journal Entry"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSchoolEquip 
         Caption         =   "&3 School Furniture and Equipment"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MnuAnlRep 
         Caption         =   "&4 Report Signing Information"
      End
      Begin VB.Menu MnuTBalInd 
         Caption         =   "&5 Trial Balance"
      End
      Begin VB.Menu MnuTBalCon 
         Caption         =   "&6 Contra Report"
      End
   End
   Begin VB.Menu MnuFSt 
      Caption         =   "&3 Financial Statements"
      Begin VB.Menu MnuIndFSt 
         Caption         =   "&1 Individual"
         Begin VB.Menu MnuBSheetTrn 
            Caption         =   "&1 Balance Sheet"
         End
         Begin VB.Menu MnuIncExpTrn 
            Caption         =   "&2 Income Expenditure Account"
         End
      End
      Begin VB.Menu MnuConstrn 
         Caption         =   "&2 Consolidated"
         Begin VB.Menu MnuBSConTrn 
            Caption         =   "&1 Balance Sheet"
         End
         Begin VB.Menu MnuIEConTrn 
            Caption         =   "&2 Income Expenditure Account"
         End
         Begin VB.Menu MnuSch9C 
            Caption         =   "&3 Schedule 9C"
         End
      End
   End
   Begin VB.Menu MnuReport 
      Caption         =   "&4 Reports"
      Begin VB.Menu MnuAuditReport 
         Caption         =   "&1 Audit Report"
      End
      Begin VB.Menu MnuFCReport 
         Caption         =   "&2 FC Report"
      End
      Begin VB.Menu MnuSchoolMemo 
         Caption         =   "&3 School Memo"
      End
      Begin VB.Menu MnuUnitReport 
         Caption         =   "&4 Unit Report"
      End
      Begin VB.Menu MnuITRcomp 
         Caption         =   "&5 Draft ITR Computation"
      End
   End
   Begin VB.Menu MnuExit 
      Caption         =   "&5 Exit"
   End
End
Attribute VB_Name = "MDIWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    If (mLVer <> mSVer) Then
        MsgBox "New Update Available. Software will update now.", vbInformation, "Update Available"
        SoftUp
    End If
End Sub

Private Sub MnuAnlRep_Click()
    PctMdi.Visible = False
    FrmAnnualRMst.Show
End Sub

Private Sub MnuAuditReport_Click()
    PctMdi.Visible = False
    FrmAuditReport.Show
End Sub

Private Sub MnuBSheetTrn_Click()
    PctMdi.Visible = False
    FrmBS.Show
End Sub

Private Sub MnuBSConTrn_Click()
    PctMdi.Visible = False
    FrmBSCon.Show
End Sub

Private Sub MnuClient_Click()
    PctMdi.Visible = False
    FrmClientMst.Show
End Sub

Private Sub MnuExit_Click()
    PctMdi.Visible = False
    FrmShut.Show
End Sub

Private Sub MnuFCReport_Click()
    PctMdi.Visible = False
    FrmFCReport.Show
End Sub

Private Sub MnuIEConTrn_Click()
    PctMdi.Visible = False
    FrmPlAcCon.Show
End Sub

Private Sub MnuIncExpTrn_Click()
    PctMdi.Visible = False
    FrmPlAc.Show
End Sub

Private Sub MnuITRcomp_Click()
    PctMdi.Visible = False
    FrmITRCom.Show
End Sub

Private Sub MnuJv_Click()
    PctMdi.Visible = False
    FrmJV.Show
End Sub

Private Sub MnuOpenBal_Click()
    PctMdi.Visible = False
    FrmUOpenBal.Show
End Sub

Private Sub MnuRecPay_Click()
    PctMdi.Visible = False
    FrmRecPay.Show
End Sub

Private Sub MnuSch9C_Click()
    PctMdi.Visible = False
    FrmSchedule9C.Show
End Sub

Private Sub MnuSchoolEquip_Click()
    PctMdi.Visible = False
    FrmFurEquip.Show
End Sub

Private Sub MnuSchoolMemo_Click()
    PctMdi.Visible = False
    FrmSchoolMemo.Show
End Sub

Private Sub MnuTBalCon_Click()
    PctMdi.Visible = False
    FrmTBalCon.Show
End Sub
Private Sub MnuTBalInd_Click()
    PctMdi.Visible = False
    FrmTBalInd.Show
End Sub
Private Sub MnuSoftUp_Click()
    If MsgBox("Are you sure to update?", vbCritical + vbYesNo, "Alert") = vbYes Then
        SoftUp
    End If
End Sub
Private Sub MnuUnitReport_Click()
    PctMdi.Visible = False
    FrmUnitRptDtl.Show
End Sub
Private Sub SoftUp()
Dim uPath As String
Dim DBLocalDB As New ADODB.Connection
Dim RsQ As New ADODB.Recordset
Set DBLocalDB = Nothing
DBLocalDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & hPath & "'"
DBLocalDB.Open
Set RsQ = Nothing
RsQ.Open "Select * From SPath", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
RsQ.MoveFirst
uPath = RsQ.Fields("SysPath")
DBLocalDB.Close
DbDataDB.Close
Set DBLocalDB = Nothing
Set DbDataDB = Nothing
Shell uPath + "\Update.bat"
End
End Sub
