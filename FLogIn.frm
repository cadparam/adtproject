VERSION 5.00
Begin VB.Form FrmLogIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FLogIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FLogIn.frx":2872A
   ScaleHeight     =   1995
   ScaleWidth      =   6120
   Begin VB.Frame FraLogin 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox ComF_Year 
         Height          =   360
         Left            =   1725
         TabIndex        =   0
         Top             =   315
         Width           =   2385
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "C&ancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         ToolTipText     =   "Exit Program"
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton CmdOk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Connect"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   2
         ToolTipText     =   "Connect to Server"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox TxtPassword 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1725
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Enter Password"
         Top             =   840
         Width           =   2385
      End
      Begin VB.Shape ShpLog 
         BackColor       =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   1695
         Left            =   120
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label LblYear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   6
         Top             =   375
         Width           =   645
      End
      Begin VB.Label LblBranch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Access Code:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   5
         Top             =   765
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBLocalDB As New ADODB.Connection
Dim RSSys As New ADODB.Recordset
Dim RSSysUser As New ADODB.Recordset
Dim DBSPath As New ADODB.Connection
Dim RsSPath As New ADODB.Recordset
Dim i As Double

Private Sub Form_Activate()
    'TxtPassword.Text = "PARAM10"
    TxtPassword.SetFocus
    'CmdOk.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub CmdCancel_Click()
    End
End Sub
Private Sub CmdOk_Click()
Dim mi As Double
If ComF_Year.Text = "" Then
    MsgBox "Please Select Financial Year!!", vbCritical, "Error"
    ComF_Year.SetFocus
    Exit Sub
End If
If TxtPassword.Text = "" Then
    MsgBox "Please Enter Password!! ", vbCritical, "LogIn Error"
    TxtPassword.SetFocus
    Exit Sub
End If
Set RSSysUser = Nothing
RSSysUser.Open "Select * From TranBr Where F_Year='" & ComF_Year.Text & "'", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
If RSSysUser.EOF = False Then
    mBYear = RSSysUser.Fields("F_Code")
    mDsnName = RSSysUser.Fields("DsnName")
    MSCONNECT = "DSN=" + mDsnName + ";USERID=ADMIN;PWD=;DBQ="
    sPath = sPath + Trim(Str(RSSysUser.Fields("Branch_Code"))) + "\" + ComF_Year.Text + "\DataDB.Mdb"
'    If MyDsnConnect.Create_Year_DSN(False, mDsnName, sPath) = False Then
'        MsgBox "Sorry! DSN Not Found !! Contact Administrator !!", vbCritical, Me.Caption
'        TxtPassword.SetFocus
'        Exit Sub
'    End If
    Set DbDataDB = Nothing
    DbDataDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & sPath     'mDsnName
'    DbDataDB.ConnectionString = "Provider=SQLOLEDB; Data Source=192.168.1.11; Initial Catalog=atbase; User ID=auditrust; Password=SQLadm722;"
    DbDataDB.Open
    OpenRecordSet
    mPassword = UCase(TxtPassword.Text)
    If RsUser.BOF = False Then
        RsUser.MoveFirst
        RsUser.Find "UPass='" & mPassword & "'"
    End If
    If RsUser.EOF = False Then
        If RsUser.Fields("UPass") = mPassword Then
            Set RsComp = Nothing
            If IIf(IsNull(RsUser.Fields("CpCode")), "", RsUser.Fields("CpCode")) <> "" Then
                RsComp.Open "Select * From FrmMst Where CpCode='" & RsUser.Fields("CpCode") & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                mCmpName = RsComp.Fields("FName") + Space(2)
            Else
                MsgBox "Sorry No Company Has Been Linked With User Account.", vbCritical, "Alert"
                RsComp.Open "Select * From FrmMst Where CpCode='AND'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            mUType = RsUser.Fields("UType")
            mSPlace = RsComp.Fields("City")
            mYear = "01-04-20" & Mid(ComF_Year.Text, 1, 2)
            mTYear = "31-03-20" & Format(Val(Mid(ComF_Year.Text, 1, 2)) + 1, "00")
            mFinYear = ComF_Year.Text
            Unload Me
            MDIWork.Show
        Else
            MsgBox "Invalid Password", vbCritical, "Login Failed"
            mUser = ""
            mPassword = ""
            sPath = hPath
            TxtPassword.SetFocus
            SendKeys "{Home}+{End}"
        End If
    Else
        MsgBox "Invalid User Name", vbCritical, "Login Failed"
        mUser = ""
        mPassword = ""
        sPath = hPath
        TxtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
    DbDataDB.BeginTrans
'    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "','LOGIN','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    MDIWork.Caption = "AudIToR 2.0 (FY 20" & CStr(mFinYear) & ")"
End If
End Sub
Private Sub Form_Load()
Dim mSysPath As String
Dim mDBPath As String
Dim DBSPath As New ADODB.Connection
Dim RsSPath As New ADODB.Recordset
Dim i As Double
hPath = App.Path + "\LocalDB.MDB"
    Set DBSPath = Nothing
    DBSPath.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & hPath & "'"
    DBSPath.Open
    Set RsSPath = Nothing
    RsSPath.Open "Select * From SPath", DBSPath, adOpenDynamic, adLockReadOnly, adCmdText
    If RsSPath.EOF = False Then
'        hPath = RsSPath.Fields("SysPath")
        sPath = RsSPath.Fields("SysPath")
        mLVer = RsSPath.Fields("Version")
'    Else
'        MsgBox "Unauthorised Access Of Programe!! " + vbCrLf + " Please Contact Administrator.", vbCritical, "Critical Error"
    End If
Set DBLocalDB = Nothing
    mDBPath = sPath + "Update\LocalDB.MDB"
    DBLocalDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & mDBPath & "'"
    DBLocalDB.Open
        Set RsSPath = Nothing
        RsSPath.Open "Select * From SPath", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        If RsSPath.EOF = False Then
            mSVer = RsSPath.Fields("Version")
        End If
    DBLocalDB.Close
Set DBLocalDB = Nothing
    DBLocalDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & hPath & "'"
    DBLocalDB.Open
    Set RSSys = Nothing
    RSSys.Open "Select * From TranBr", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    RSSys.MoveFirst
    ComF
    mUser = ""
    mPassword = ""
    Me.Top = 3810
    Me.Left = 3375
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set RSSys = Nothing
End Sub
Private Function ComF()
Dim RsTran_Qry As New ADODB.Recordset
RsTran_Qry.Open "Select F_Year From TranBr Order By F_Code Desc", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsTran_Qry.EOF = False Then
    RsTran_Qry.MoveFirst
    ComF_Year.Clear
    Do While RsTran_Qry.EOF = False
        ComF_Year.AddItem RsTran_Qry.Fields("F_Year")
        RsTran_Qry.MoveNext
    Loop
    ComF_Year.ListIndex = 0
End If
End Function
