VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmAnnualRMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Signing Information"
   ClientHeight    =   8790
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
   Icon            =   "FAnlRepM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   15705
   Begin VB.Frame FraHelp 
      Height          =   6732
      Left            =   18000
      TabIndex        =   37
      Top             =   0
      Width           =   13092
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         Height          =   384
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   6240
         Visible         =   0   'False
         Width           =   1056
      End
      Begin VB.CommandButton CmdLClose 
         Caption         =   "Cl&ose"
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit"
         Top             =   6240
         Width           =   850
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfUDIN 
         Height          =   5892
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   12852
         _cx             =   22669
         _cy             =   10393
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
         BackColorBkg    =   16777088
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
         FormatString    =   $"FAnlRepM.frx":0442
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
         Height          =   6612
         Index           =   2
         Left            =   0
         Top             =   120
         Width           =   13092
      End
   End
   Begin VB.Frame FraAcHelp 
      Height          =   6612
      Left            =   18000
      TabIndex        =   33
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
         TabIndex        =   35
         ToolTipText     =   "Exit"
         Top             =   6120
         Width           =   972
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfAcHelp 
         Height          =   5772
         Left            =   120
         TabIndex        =   34
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
         FormatString    =   $"FAnlRepM.frx":048B
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
      TabIndex        =   21
      Top             =   1080
      Width           =   15492
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
         Left            =   13920
         Picture         =   "FAnlRepM.frx":04D4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add"
         Top             =   1200
         Width           =   372
      End
      Begin VB.ComboBox TxtSAuditNo 
         Height          =   324
         Left            =   9480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   1320
      End
      Begin VB.TextBox TxtRUdIN 
         Height          =   384
         Left            =   11160
         TabIndex        =   9
         Top             =   1200
         Width           =   2736
      End
      Begin VB.TextBox TxtSTrustee 
         Height          =   384
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   4416
      End
      Begin VB.TextBox TxtSignPlace 
         Height          =   384
         Left            =   1920
         TabIndex        =   7
         Top             =   1200
         Width           =   1656
      End
      Begin VB.TextBox TxtAName 
         Height          =   384
         Left            =   12480
         TabIndex        =   13
         Top             =   720
         Width           =   2856
      End
      Begin VB.TextBox TxtName 
         Height          =   360
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   5330
      End
      Begin VB.TextBox TxtFileNo 
         Height          =   384
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1656
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   6012
         Left            =   120
         TabIndex        =   0
         Top             =   1596
         Width           =   15252
         _cx             =   26903
         _cy             =   10604
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
         FormatString    =   $"FAnlRepM.frx":0B0A
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
      Begin MSMask.MaskEdBox DtpFDate 
         Height          =   345
         Left            =   9360
         TabIndex        =   3
         ToolTipText     =   "Enter From Date"
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DtpTDate 
         Height          =   345
         Left            =   11520
         TabIndex        =   4
         ToolTipText     =   "Enter From Date"
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DtpRDate 
         Height          =   345
         Left            =   7800
         TabIndex        =   8
         ToolTipText     =   "Enter From Date"
         Top             =   1200
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label LblCompany 
         Caption         =   "Report UDIN :"
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
         Index           =   10
         Left            =   9720
         TabIndex        =   31
         Top             =   1200
         Width           =   1536
      End
      Begin VB.Label LblCompany 
         Caption         =   "Report Date :"
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
         Index           =   9
         Left            =   6480
         TabIndex        =   30
         Top             =   1200
         Width           =   1476
      End
      Begin VB.Label LblCompany 
         Caption         =   "Signing Trustee :"
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
         TabIndex        =   29
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label LblCompany 
         Caption         =   "Signing Place :"
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
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1452
      End
      Begin VB.Label LblCompany 
         Caption         =   "Auditor Name :"
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
         Index           =   6
         Left            =   10800
         TabIndex        =   27
         Top             =   720
         Width           =   1656
      End
      Begin VB.Label LblCompany 
         Caption         =   "Signing Auditor (Mem. No.) :"
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
         Left            =   6480
         TabIndex        =   26
         Top             =   720
         Width           =   2880
      End
      Begin VB.Label LblCompany 
         Caption         =   "To :"
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
         Left            =   10920
         TabIndex        =   25
         Top             =   240
         Width           =   480
      End
      Begin VB.Label LblCompany 
         Caption         =   "Year From :"
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
         Left            =   8040
         TabIndex        =   24
         Top             =   240
         Width           =   1212
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
         TabIndex        =   22
         Top             =   240
         Width           =   912
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   7572
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   15492
      End
   End
   Begin VB.Frame FraTool 
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   7092
      Begin VB.CommandButton TlbSav 
         Caption         =   "Pr&int"
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
         Picture         =   "FAnlRepM.frx":0B53
         Style           =   1  'Graphical
         TabIndex        =   36
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
         Picture         =   "FAnlRepM.frx":11BD
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "FAnlRepM.frx":15FF
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "FAnlRepM.frx":1A41
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "FAnlRepM.frx":20AB
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "FAnlRepM.frx":24ED
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "FAnlRepM.frx":292F
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
      TabIndex        =   32
      Top             =   480
      Width           =   4956
   End
End
Attribute VB_Name = "FrmAnnualRMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim mActivity As String
Dim mAcCode As Double
Dim RsMember As New ADODB.Recordset
Dim mAcList As String
Private Sub CmdAClose_Click()
    FraAcHelp.Left = 18000
    TxtFileNo.SetFocus
End Sub
Private Sub CmdLClose_Click()
    FraHelp.Left = 15600
End Sub
Private Sub CmdSearch_Click()
    FraHelp.Left = 2000
    SetParent
    Dim RsQ As New ADODB.Recordset
    DBWorkTmp.BeginTrans
    DBWorkTmp.Execute "Delete From TmpTrialBal"
    DBWorkTmp.CommitTrans
    DBWorkTmp.BeginTrans
    RsQ.Open "Select * From QTrialBal Where AcCode In (" & mAcList & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpTrialBal (HType,HSide,AcCode,HCode,LCode,OpDr,OpCr,ADr,ACr,DBal,CBal) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & _
        "'," & RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("OpDr") & "," & RsQ.Fields("OpCr") & "," & RsQ.Fields("ADr") & _
        "," & RsQ.Fields("ACr") & "," & RsQ.Fields("DBal") & "," & RsQ.Fields("CBal") & ")"
        RsQ.MoveNext
    Loop
    DBWorkTmp.CommitTrans
    SetData
End Sub
Private Sub DtpFDate_Validate(Cancel As Boolean)
If Len(Trim(DtpFDate.Text)) < 10 Then
    MsgBox "Invalid Date.", vbInformation, "Alert"
    DtpFDate.Text = "  -  -    "
    DtpFDate.SetFocus
End If
End Sub
Private Sub DtptDate_Validate(Cancel As Boolean)
If Len(Trim(DtpTDate.Text)) < 10 Then
    MsgBox "Invalid Date.", vbInformation, "Alert"
    DtpTDate.Text = "  -  -    "
    DtpTDate.SetFocus
End If
End Sub
Private Sub DtpRDate_Validate(Cancel As Boolean)
If Len(Trim(DtpRDate.Text)) < 10 Then
    MsgBox "Invalid Date.", vbInformation, "Alert"
    DtpRDate.Text = "  -  -    "
    DtpRDate.SetFocus
End If
End Sub

Private Sub Form_Load()
    DBWorkTmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & App.Path + "\LocalDB.Mdb'"
    DBWorkTmp.Open
    Me.Left = 50
    Me.Top = 50
    SetCombo
    If VsfHelp.Rows > 1 Then Display
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.FileNo,AcMst.AcName,RepDtl.*,AcMst.AcType,AdtMst.AdtName From AcMst,RepDtl,AdtMst Where AcMst.Active=-1 And AcMst.AcCode=RepDtl.AcCode And AdtMst.AdtNo=RepDtl.AdtNo Order By FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfHelp.DataSource = RsQry
With VsfHelp
    .TextMatrix(0, 0) = "FILE NO"
    .ColWidth(0) = 1000
    .TextMatrix(0, 1) = "CLIENT NAME"
    .ColWidth(1) = 3500
    .ColWidth(2) = 0    'AcCode
    .TextMatrix(0, 3) = "FROM"
    .ColWidth(3) = 1100
    .TextMatrix(0, 4) = "TO"
    .ColWidth(4) = 1100
    .TextMatrix(0, 5) = "REPORT DT"
    .ColWidth(5) = 1100
    .TextMatrix(0, 6) = "PLACE"
    .ColWidth(6) = 800
    .TextMatrix(0, 7) = "TRUSTEE"
    .ColWidth(7) = 1000
    .TextMatrix(0, 9) = "AUDIT CD"
    .ColWidth(9) = 1000
    .TextMatrix(0, 8) = "UDIN"
    .ColWidth(8) = 2500
    .ColWidth(10) = 0    'TYPECODE
    .TextMatrix(0, 11) = "AUDITOR"
    .ColWidth(11) = 1500
    .Refresh
End With
Set RsQry = Nothing
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,GrpMst.GName,AcMst.AcCode From AcMst,GrpMst Where AcMst.AcType=" & _
"GrpMst.GCode And AcMst.Active=-1 Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfAcHelp.DataSource = RsQry
With VsfAcHelp
    .TextMatrix(0, 0) = "NAME"
    .ColWidth(0) = 4800
    .TextMatrix(0, 1) = "FILE NO."
    .ColWidth(1) = 1500
    .TextMatrix(0, 2) = "CITY"
    .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "TYPE"
    .ColWidth(3) = 1000
    .TextMatrix(0, 4) = "ACCODE"
    .ColWidth(4) = 0
    .Col = 1
    .ExtendLastCol = True
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
End With
Set RsMember = Nothing
RsMember.Open "Select * From AdtMst Order By AdtName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
TxtSAuditNo.Clear
Do While RsMember.EOF = False
    TxtSAuditNo.AddItem RsMember.Fields("AdtNo")
    TxtSAuditNo.ItemData(TxtSAuditNo.NewIndex) = RsMember.Fields("AdtNo")
    RsMember.MoveNext
Loop
TxtSAuditNo.ListIndex = 0
End Function
Private Sub ClearText()
Dim ObjText As Object
For Each ObjText In Me
    If TypeOf ObjText Is TextBox Then ObjText.Text = ""
    If TypeOf ObjText Is MaskEdBox Then ObjText.Text = "  -  -    "
Next
End Sub
Private Function SaveData()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From RepDtl Where AcCode=" & mAcCode
If mActivity <> "Delete" Then
    DbDataDB.Execute "Insert InTo RepDtl (AcCode,RFDt,RTDt,AdtNo,RPlace,RTrustee,RPDt,RUDIN) Values (" & mAcCode & ",'" & DtpFDate.Text & "','" & DtpTDate.Text & "','" & TxtSAuditNo.Text & "','" & _
    TxtSignPlace.Text & "','" & TxtSTrustee.Text & "','" & DtpRDate.Text & "','" & TxtRUdIN.Text & "')"
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

Private Sub TxtRUdIN_Validate(Cancel As Boolean)
If TxtRUdIN.Text <> "" Then
    If Len(Trim(TxtRUdIN.Text)) <> 18 Then
        MsgBox "Invalid UDIN.", vbInformation + vbCritical, "Alert"
        TxtRUdIN.Text = ""
    ElseIf CheckUDIN(Mid(TxtRUdIN.Text, 1, 8), "N") = True Then
        MsgBox "Invalid UDIN.", vbInformation + vbCritical, "Alert"
        'TxtRUdIN.Text = ""
    ElseIf CheckUDIN(Mid(TxtRUdIN.Text, 9, 6), "C") = True Then
        MsgBox "Invalid UDIN.", vbInformation + vbCritical, "Alert"
        'TxtRUdIN.Text = ""
    ElseIf CheckUDIN(Mid(TxtRUdIN.Text, 15, 4), "N") = True Then
        MsgBox "Invalid UDIN.", vbInformation + vbCritical, "Alert"
        'TxtRUdIN.Text = ""
    End If
End If
End Sub

Private Sub TxtSAuditNo_Validate(Cancel As Boolean)
If TxtSAuditNo.Text <> "" Then
    If RsMember.BOF = False Then
        RsMember.MoveFirst
        RsMember.Find "AdtNo='" & TxtSAuditNo.Text & "'"
    End If
    If RsMember.EOF = False Then
        TxtAName.Text = RsMember.Fields("AdtName")
    Else
        RsMember.MoveFirst
        TxtAName.Text = RsMember.Fields("AdtName")
    End If
End If
End Sub

Private Sub VsfHelp_EnterCell()
    If mActivity = "Add" Or mActivity = "Edit" Then
        If VsfHelp.Col = 0 Or VsfHelp.Col = 2 Then VsfHelp.Editable = flexEDKbd Else VsfHelp.Editable = flexEDNone
    End If
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
        ElseIf Len(Trim(DtpFDate.Text)) < 10 Then
            MsgBox "Sorry! Invalid Date.", vbInformation, "Black Data Error"
            DtpFDate.SetFocus
        ElseIf Len(Trim(DtpTDate.Text)) < 10 Then
            MsgBox "Sorry! Invalid Date.", vbInformation, "Black Data Error"
            DtpTDate.SetFocus
        ElseIf Len(Trim(DtpRDate.Text)) < 10 Then
            MsgBox "Sorry! Invalid Date.", vbInformation, "Black Data Error"
            DtpRDate.SetFocus
        ElseIf DupliRec = True Then
            MsgBox "Sorry! Duplicate Record Found.", vbInformation, "Black Data Error"
            TxtFileNo.SetFocus
        ElseIf DupliUDIN = True Then
            MsgBox "Sorry! Duplicate Record Found.", vbInformation, "Black Data Error"
            TxtRUdIN.SetFocus
        Else
            SaveData
            DbDataDB.BeginTrans
                If mActivity = "Add" Then
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SIGN_INFO','ADD_NEW','" & Date & "','" & Time & "')"
                ElseIf mActivity = "Delete" Then
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SIGN_INFO','DELETE','" & Date & "','" & Time & "')"
                ElseIf mActivity = "Edit" Then
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SIGN_INFO','UPDATE','" & Date & "','" & Time & "')"
                End If
            DbDataDB.CommitTrans
            mActivity = ""
            ClearText
            FraHelp.Left = 15600
            VsfHelp.Enabled = True
            SetCombo
            SetTool True
            If VsfHelp.Rows > 1 Then Display
            VsfHelp.SetFocus
        End If
    Case "Cancel"
        mActivity = ""
        VsfHelp.Enabled = True
        FraHelp.Left = 15600
        SetTool True
        ClearText
        SetCombo
        SetTool True
        If VsfHelp.Rows > 1 Then Display
        VsfHelp.SetFocus
    Case "Exit"
        DBWorkTmp.Close
        RsMember.Close
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
Dim RsQ As New ADODB.Recordset
With VsfHelp
     TxtFileNo.Text = .TextMatrix(.Row, 0)
     TxtName.Text = .TextMatrix(.Row, 1)
     mAcCode = Val(.TextMatrix(.Row, 2))
     DtpFDate.Text = Format(IIf(IsNull(.TextMatrix(.Row, 3)) = False, .TextMatrix(.Row, 3), mYear), "dd-MM-yyyy")
     DtpTDate.Text = Format(IIf(IsNull(.TextMatrix(.Row, 4)) = False, .TextMatrix(.Row, 4), mTYear), "dd-MM-yyyy")
     DtpRDate.Text = Format(.TextMatrix(.Row, 5), "dd-MM-yyyy")
     TxtSignPlace.Text = IIf(IsNull(.TextMatrix(.Row, 6)) = False, .TextMatrix(.Row, 6), mSPlace)
     TxtSTrustee.Text = .TextMatrix(.Row, 7)
     If .TextMatrix(.Row, 9) <> "" Then TxtSAuditNo.Text = .TextMatrix(.Row, 9)
     TxtRUdIN.Text = .TextMatrix(.Row, 8)
     TxtAName.Text = .TextMatrix(.Row, 11)
End With
End Sub
Private Sub VsfAcHelp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TxtName.Text = VsfAcHelp.TextMatrix(VsfAcHelp.Row, 0)
        TxtFileNo.Text = VsfAcHelp.TextMatrix(VsfAcHelp.Row, 1)
        mAcCode = Val(VsfAcHelp.TextMatrix(VsfAcHelp.Row, 4))
        DtpFDate.Text = Format(mYear, "dd-MM-yyyy")
        DtpTDate.Text = Format(mTYear, "dd-MM-yyyy")
        TxtSignPlace.Text = mSPlace
        FraAcHelp.Left = 18000
        TxtFileNo.SetFocus
    End If
End Sub

Private Sub SetData()
Dim RsQ As New ADODB.Recordset
Dim RsQParent As New ADODB.Recordset
Dim mPaCode As Double
RsQ.Open "Select * From AcMst Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With VsfUDIN
    .Clear
    .Cols = 3
    .Rows = 1
    .Rows = 13
    .FixedCols = 1
    .TextMatrix(0, 0) = "SR"
    .ColWidth(0) = 500
    .ColAlignment(0) = flexAlignRightCenter
    .TextMatrix(0, 1) = "PARTICULAR"
    .ColWidth(1) = 4500
    .ColAlignment(1) = flexAlignLeftCenter
    .TextMatrix(0, 2) = "VALUE"
    .ColWidth(2) = 7500
    .ColAlignment(2) = flexAlignLeftCenter
    .Row = 1
    .Col = 1
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "FRN"
    .TextMatrix(.Row, 2) = RsComp.Fields("FRN")
    .Row = .Row + 1 '   2
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "Document Type"
    .TextMatrix(.Row, 2) = "Audit & Assurance Functions"
    .Row = .Row + 1 '   3
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "Type of Audit"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("AcType") = 1 Or RsQ.Fields("AcType") = 2, "Statutory Audit - Non-Corporate", "Income/Receipt and Payment/Expenditure Audit")
    .Row = .Row + 1 '   4
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "Type of Audit"
    If RsQ.Fields("AcType") = 1 Then
        .TextMatrix(.Row, 2) = "Maharashtra Public Trusts Act, 1950"
    ElseIf RsQ.Fields("AcType") = 2 Then
        .TextMatrix(.Row, 2) = "Foreign Contribution (Regulation) Act, 2010"
    ElseIf RsQ.Fields("AcType") = 3 Then
        .TextMatrix(.Row, 2) = "Other Acts/Regulations/Laws/Statutes not covered above"
    Else
        .TextMatrix(.Row, 2) = "Not Applicable under Any Act/Regulation/Statute"
    End If
    .Row = .Row + 1 '   5
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "Date of Signing Document"
    .TextMatrix(.Row, 2) = DtpRDate.Text
    .Row = .Row + 1 '   6
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "Financial Year"
    .TextMatrix(.Row, 2) = DtpFDate.Text + "-" + DtpTDate.Text
    .Row = .Row + 1 '   7
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "PAN of the Assessee/Auditee"
    If RsQ.Fields("TPan") <> "" Then
        .TextMatrix(.Row, 2) = RsQ.Fields("TPan")
    ElseIf IsNull(RsQ.Fields("PaCode")) = False Then
        Set RsQ = Nothing
        RsQ.Open "Select SaCode from QGroup where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        mPaCode = RsQ.Fields("SaCode")
        Set RsQ = Nothing
        RsQ.Open "Select * From AcMst Where AcCode=" & mPaCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQ.EOF = False Then .TextMatrix(.Row, 2) = IIf(IsNull(RsQ.Fields("TPan")) = True, "", RsQ.Fields("TPan"))
        Set RsQ = Nothing
        RsQ.Open "Select * From AcMst Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    End If
    .Row = .Row + 1 '   8
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "Gross Turnover/Gross Receipt/Gross Income"
    SetGrossInc
    .TextMatrix(.Row, 2) = TxtTotal.Text
    .Row = .Row + 1 '   9
    .TextMatrix(.Row, 0) = .Row
    If RsQ.Fields("AcType") = 1 Or RsQ.Fields("AcType") = 2 Then
        .TextMatrix(.Row, 1) = "Shareholders Fund/Owners Fund"
        SetShareFund
        .TextMatrix(.Row, 2) = TxtTotal.Text
    Else
        .TextMatrix(.Row, 1) = "Any Comment/Recommendation/Adverse Comment"
        .TextMatrix(.Row, 2) = "No"
    End If
    .Row = .Row + 1 '   10
    .TextMatrix(.Row, 0) = .Row
    If RsQ.Fields("AcType") = 1 Or RsQ.Fields("AcType") = 2 Then
        .TextMatrix(.Row, 1) = "Net Block of Property, Plant & Equipment"
        SetPlant
        .TextMatrix(.Row, 2) = TxtTotal.Text
    Else
        .TextMatrix(.Row, 1) = "Any Comment/Recommendation/Adverse Comment"
        .TextMatrix(.Row, 2) = "No"
    End If
    .Row = .Row + 1 '   11
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "Document Description"
    If RsQ.Fields("AcType") = 1 Then
        .TextMatrix(.Row, 2) = "Independent Auditor`s Report"
    ElseIf RsQ.Fields("AcType") = 2 Then
        .TextMatrix(.Row, 2) = "Certificate of Chartered Accountant under FC(R) Act"
    ElseIf RsQ.Fields("AcType") = 3 Then
        .TextMatrix(.Row, 2) = "Audited Memo of Receipts and Expenditure of School"
    ElseIf RsQ.Fields("AcType") = 4 Then
        .TextMatrix(.Row, 2) = "Audited Financial Statements of Branch of Trust"
    ElseIf RsQ.Fields("AcType") = 5 Then
        .TextMatrix(.Row, 2) = "Audited Income and Expenditure Account for Government Grant"
    End If
    .Row = .Row + 1 '   12
    .TextMatrix(.Row, 0) = .Row
    .TextMatrix(.Row, 1) = "Remarks"
    If RsQ.Fields("AcType") = 2 Then
        If IsNull(RsQ.Fields("PaCode")) = False Then
            mPaCode = RsQ.Fields("PaCode")
            Set RsQ = Nothing
            RsQ.Open "Select * From AcMst Where AcCode=" & mPaCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsQ.EOF = False Then .TextMatrix(.Row, 2) = RsQ.Fields("AcName") & IIf(RsQ.Fields("City") <> "", ", " & RsQ.Fields("City"), "")
            Set RsQ = Nothing
            RsQ.Open "Select * From AcMst Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        End If
    Else
        .TextMatrix(.Row, 2) = RsQ.Fields("AcName") & IIf(RsQ.Fields("City") <> "", ", " & RsQ.Fields("City"), "")
    End If
    .Editable = flexEDKbd
    .Refresh
End With
End Sub
Private Sub SetParent()
Dim RsQ As New ADODB.Recordset
If Val(VsfHelp.TextMatrix(VsfHelp.Row, 10)) = 2 Then
    Set RsQ = Nothing
    RsQ.Open "Select * From QGroup Where PaCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    mAcList = CStr(mAcCode)
    Do While RsQ.EOF = False
        mAcList = mAcList + "," + CStr(RsQ.Fields("AcCode"))
        RsQ.MoveNext
    Loop
ElseIf Val(VsfHelp.TextMatrix(VsfHelp.Row, 10)) = 1 Then
    Set RsQ = Nothing
    RsQ.Open "Select * From QGroup Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("SACode") = mAcCode Then
        Set RsQ = Nothing
        RsQ.Open "Select * From QGroup Where SaCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQ.EOF = False
            If mAcList = "" Then mAcList = CStr(RsQ.Fields("AcCode")) Else mAcList = mAcList + "," + CStr(RsQ.Fields("AcCode"))
            RsQ.MoveNext
        Loop
    End If
Else
    mAcList = mAcCode
End If
End Sub
Private Sub SetGrossInc()
Dim RsQry As New ADODB.Recordset
TxtTotal.Text = "0.00"
If Val(VsfHelp.TextMatrix(VsfHelp.Row, 10)) <= 2 Then
    Set RsQry = Nothing
    RsQry.Open "Select Sum(CBal) As TotRs From TmpTrialBal Where HType=0 And LCode<>108 And HType<>54", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    TxtTotal.Text = IIf(IsNull(RsQry.Fields("TotRs")) = True, 0, RsQry.Fields("TotRs"))
Else
    Set RsQry = Nothing
    RsQry.Open "Select Sum(CBal) As TotRs From TmpTrialBal Where HType=0 And HType<>54", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    TxtTotal.Text = IIf(IsNull(RsQry.Fields("TotRs")) = True, 0, RsQry.Fields("TotRs"))
End If
End Sub
Private Sub SetShareFund()
Dim RsQry As New ADODB.Recordset
TxtTotal.Text = "0.00"
If Val(VsfHelp.TextMatrix(VsfHelp.Row, 10)) <= 2 Then
    Set RsQry = Nothing
    RsQry.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where HCode In (58,57,5,2,1)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    TxtTotal.Text = RsQry.Fields("TotRs") + SetProfitAll(mAcList)
Else
    Set RsQry = Nothing
    RsQry.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode=" & mAcCode & " And HCode In (58,57,5,2,1)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    TxtTotal.Text = RsQry.Fields("TotRs") + SetProfit(mAcCode)
End If
End Sub
Private Sub SetPlant()
Dim RsQry As New ADODB.Recordset
TxtTotal.Text = "0.00"
Set RsQry = Nothing
RsQry.Open "Select Sum(IIF(IsNull(DBal)=True,0,DBal)) As TotRs From TmpTrialBal Where HCode In (6,8)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
TxtTotal.Text = IIf(IsNull(RsQry.Fields("TotRs")) = False, RsQry.Fields("TotRs"), "0")
End Sub

Private Function DupliRec() As Boolean
Dim RsQ As New ADODB.Recordset
If mAcCode <> Val(VsfHelp.TextMatrix(VsfHelp.Row, 2)) Then
    RsQ.Open "Select * From RepDtl Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.EOF = False Then DupliRec = True Else DupliRec = False
Else
    DupliRec = False
End If
End Function
Private Function DupliUDIN() As Boolean
Dim RsQ As New ADODB.Recordset
If TxtRUdIN.Text <> "" Then
    If mAcCode <> Val(VsfHelp.TextMatrix(VsfHelp.Row, 2)) Then
        RsQ.Open "Select * From RepDtl Where RUDIN='" & TxtRUdIN.Text & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQ.EOF = False Then DupliUDIN = True Else DupliUDIN = False
    Else
        If TxtRUdIN.Text <> VsfHelp.TextMatrix(VsfHelp.Row, 8) Then
            RsQ.Open "Select * From RepDtl Where RUDIN='" & TxtRUdIN.Text & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsQ.EOF = False Then DupliUDIN = True Else DupliUDIN = False
        Else
            DupliUDIN = False
        End If
    End If
Else
    DupliUDIN = False
End If
End Function

Private Function CheckUDIN(ByVal mNumber As String, ByVal mCheck As String) As Boolean
Dim i As Integer
i = 1
If mCheck = "C" Then
    Do While i <= Len(mNumber)
        If IsNumeric(Mid(mNumber, i, 1)) = True Then
            CheckUDIN = True
            Exit Do
        Else
            i = i + 1
        End If
    Loop
Else
    Do While i <= Len(mNumber)
        If IsNumeric(Mid(mNumber, i, 1)) = False Then
            CheckUDIN = True
            Exit Do
        Else
            i = i + 1
        End If
    Loop
End If
End Function
