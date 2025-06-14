VERSION 5.00
Begin VB.Form FrmShut 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shut Down Windows"
   ClientHeight    =   1545
   ClientLeft      =   3045
   ClientTop       =   3285
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FShut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraShut 
      Height          =   1452
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00C0E0FF&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image ImgComputer 
         Height          =   495
         Left            =   240
         Picture         =   "FShut.frx":000C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "What do you want  the computer to do ?"
         Height          =   312
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   3252
      End
   End
End
Attribute VB_Name = "FrmShut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    MDIWork.PctMdi.Visible = True
    Unload Me
End Sub
Private Sub CmdOk_Click()
    Set RsUser = Nothing
    Set RsComp = Nothing
    Set RsGroup = Nothing
    DbDataDB.BeginTrans
    DbDataDB.Execute "Delete From TmpBSPrn"
    DbDataDB.Execute "Delete From TmpCtDtl"
    DbDataDB.Execute "Delete From TmpNotePrn"
    DbDataDB.Execute "Delete From TmpPLPrn"
    DbDataDB.Execute "Delete From TmpRecPrn"
    DbDataDB.Execute "Delete From TmpRPDtl"
    DbDataDB.Execute "Delete From TmpSMPrn"
    DbDataDB.Execute "Delete From TmpTrialBal"
'    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "','LOGOUT','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    DbDataDB.Close
    End
End Sub
