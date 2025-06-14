VERSION 5.00
Begin VB.Form FrmSplace 
   BackColor       =   &H00B4670E&
   BorderStyle     =   0  'None
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   FillColor       =   &H00B4670E&
   ForeColor       =   &H00B4670E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "FSplace.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "FrmSplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBSPath As New ADODB.Connection
Dim RsSPath As New ADODB.Recordset
Dim i As Double
Private Sub Form_Load()
    i = 1
    GetPath
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FrmLogIn.Show
End Sub
Private Sub Timer1_Timer()
    i = i + 1
    If i = 2 Then
        Form_Click
        Timer1.Enabled = False
           Unload Me
    End If
End Sub
Private Function GetPath()
Dim mSysPath As String
hPath = App.Path + "\LocalDB.MDB"
Set DBSPath = Nothing
DBSPath.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & hPath & "'"
DBSPath.Open
Set RsSPath = Nothing
RsSPath.Open "Select * From SPath", DBSPath, adOpenDynamic, adLockReadOnly, adCmdText
If RsSPath.EOF = False Then
    hPath = RsSPath.Fields("SysPath")
    sPath = RsSPath.Fields("SysPath")
Else
    MsgBox "Unautorised Access Of Programe !! " + vbCrLf + " Please  Contact Software Vendor", vbCritical, "Critical Error"
End If
Set RsSPath = Nothing
Set DBSPath = Nothing
End Function
Sub Form_Click()
   Dim CX, CY, Msg, XPos, YPos   ' Declare variables.
   ScaleMode = 3   ' Set ScaleMode to
         ' pixels.
   DrawWidth = 5   ' Set DrawWidth.
   ForeColor = QBColor(4)   ' Set foreground to red.
   FontSize = 24   ' Set point size.
   CX = ScaleWidth / 2   ' Get horizontal center.
   CY = ScaleHeight / 2   ' Get vertical center.
   Cls   ' Clear form.
   CurrentX = CX - TextWidth(Msg) / 2   ' Horizontal position.
   CurrentY = CY - TextHeight(Msg)   ' Vertical position.
   Do
    XPos = Rnd * ScaleWidth   ' Get horizontal position.
    YPos = Rnd * ScaleHeight   ' Get vertical position.
    PSet (XPos, YPos), QBColor(Rnd * 15)   ' Draw confetti.
    i = i + 1
    If i < 8000 Then
'        DoEvents
    Else
        Unload Me
        Exit Sub
    End If
   Loop   ' processing.
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'FrmLogIn.Show        ' Yield to other
End Sub
