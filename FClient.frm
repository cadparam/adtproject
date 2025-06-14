VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmClientMst 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLIENT LIST"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15630
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FClient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   15630
   Begin VB.Frame FraTool 
      Height          =   1092
      Left            =   2760
      TabIndex        =   32
      Top             =   0
      Width           =   6612
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
         Picture         =   "FClient.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Add"
         Top             =   240
         Width           =   972
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
         Left            =   1200
         Picture         =   "FClient.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Edit"
         Top             =   240
         Width           =   972
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
         Left            =   2280
         Picture         =   "FClient.frx":0EEE
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Delete"
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
         Left            =   3360
         Picture         =   "FClient.frx":1330
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   975
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
         Left            =   4440
         Picture         =   "FClient.frx":199A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancel"
         Top             =   240
         Width           =   975
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
         Left            =   5520
         Picture         =   "FClient.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   972
      End
      Begin VB.Shape ShpMain 
         Height          =   972
         Index           =   0
         Left            =   -2280
         Top             =   120
         Width           =   8892
      End
   End
   Begin VB.Frame FraCompany 
      ForeColor       =   &H00000000&
      Height          =   5532
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   15492
      Begin VSFlex7Ctl.VSFlexGrid VsfAcHelp 
         Height          =   3012
         Left            =   18000
         TabIndex        =   61
         Top             =   360
         Width           =   13452
         _cx             =   23728
         _cy             =   5313
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
         FormatString    =   $"FClient.frx":221E
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
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   5172
         Left            =   8760
         TabIndex        =   0
         Top             =   240
         Width           =   6612
         _cx             =   11663
         _cy             =   9123
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
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FClient.frx":2267
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   3495
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6165
         _Version        =   393216
         Tabs            =   5
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   732
         BackColor       =   -2147483643
         TabCaption(0)   =   "Primary Detail"
         TabPicture(0)   =   "FClient.frx":22B0
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FraPrimary"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Contact Detail"
         TabPicture(1)   =   "FClient.frx":22CC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FraContact"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Legal Details"
         TabPicture(2)   =   "FClient.frx":22E8
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "FraLegal"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Bank Details"
         TabPicture(3)   =   "FClient.frx":2304
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FraBank"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "FCRA Details"
         TabPicture(4)   =   "FClient.frx":2320
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "FraFcra"
         Tab(4).ControlCount=   1
         Begin VB.Frame FraPrimary 
            Height          =   2772
            Left            =   -74880
            TabIndex        =   55
            Top             =   528
            Width           =   6372
            Begin VB.TextBox TxtPFileNo 
               Height          =   384
               Left            =   1350
               TabIndex        =   5
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox TxtFileNo 
               Height          =   384
               Left            =   1020
               TabIndex        =   2
               Top             =   240
               Width           =   1932
            End
            Begin VB.TextBox TxtName 
               Height          =   360
               Left            =   1020
               TabIndex        =   3
               Top             =   720
               Width           =   5330
            End
            Begin VB.ComboBox ComGroup 
               Height          =   360
               Left            =   1020
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   1200
               Width           =   1455
            End
            Begin VB.TextBox TxtParentName 
               Height          =   360
               Left            =   2500
               Locked          =   -1  'True
               TabIndex        =   56
               ToolTipText     =   "Enter Phone No."
               Top             =   2160
               Width           =   3800
            End
            Begin VB.TextBox TxtParentCd 
               Height          =   360
               Left            =   9540
               Locked          =   -1  'True
               TabIndex        =   6
               ToolTipText     =   "Enter Phone No."
               Top             =   2160
               Visible         =   0   'False
               Width           =   1092
            End
            Begin VB.Shape ShpMst 
               Height          =   2652
               Index           =   1
               Left            =   0
               Top             =   120
               Width           =   6372
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
               TabIndex        =   60
               Top             =   276
               Width           =   792
            End
            Begin VB.Label LblCompany 
               Caption         =   "Name :"
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
               TabIndex        =   59
               Top             =   720
               Width           =   648
            End
            Begin VB.Label LblCompany 
               Caption         =   "Type :"
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
               Index           =   13
               Left            =   120
               TabIndex        =   58
               Top             =   1200
               Width           =   570
            End
            Begin VB.Label LblCompany 
               Caption         =   "Parent Code :"
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
               Index           =   22
               Left            =   120
               TabIndex        =   57
               Top             =   2160
               Width           =   1224
            End
         End
         Begin VB.Frame FraContact 
            Height          =   2772
            Left            =   -74760
            TabIndex        =   48
            Top             =   528
            Width           =   7212
            Begin VB.TextBox TxtAdd 
               Height          =   360
               Left            =   996
               TabIndex        =   7
               ToolTipText     =   "Enter Address"
               Top             =   240
               Width           =   6156
            End
            Begin VB.TextBox TxtCity 
               Height          =   360
               Left            =   996
               TabIndex        =   8
               ToolTipText     =   "Enter City - PinCode"
               Top             =   720
               Width           =   2652
            End
            Begin VB.TextBox TxtTaluka 
               Height          =   360
               Left            =   960
               TabIndex        =   10
               ToolTipText     =   "Enter Address"
               Top             =   1200
               Width           =   1932
            End
            Begin VB.TextBox TxtDistrict 
               Height          =   360
               Left            =   3852
               TabIndex        =   11
               ToolTipText     =   "Enter Address"
               Top             =   1200
               Width           =   1812
            End
            Begin VB.TextBox TxtState 
               Height          =   360
               Left            =   960
               TabIndex        =   12
               ToolTipText     =   "Enter State"
               Top             =   1680
               Width           =   2055
            End
            Begin VB.TextBox TxtPinCode 
               Height          =   360
               Left            =   5160
               TabIndex        =   9
               ToolTipText     =   "Enter State"
               Top             =   720
               Width           =   1935
            End
            Begin VB.Shape ShpMst 
               Height          =   2652
               Index           =   2
               Left            =   0
               Top             =   120
               Width           =   7212
            End
            Begin VB.Label LblCompany 
               Caption         =   "Address"
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
               Index           =   9
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   828
            End
            Begin VB.Label LblCompany 
               Caption         =   "City"
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
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   53
               Top             =   720
               Width           =   525
            End
            Begin VB.Label LblCompany 
               Caption         =   "Taluka :"
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
               Index           =   15
               Left            =   120
               TabIndex        =   52
               Top             =   1200
               Width           =   720
            End
            Begin VB.Label LblCompany 
               Caption         =   "District :"
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
               Index           =   16
               Left            =   3000
               TabIndex        =   51
               Top             =   1200
               Width           =   828
            End
            Begin VB.Label LblCompany 
               Caption         =   "State :"
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
               Index           =   4
               Left            =   120
               TabIndex        =   50
               Top             =   1680
               Width           =   564
            End
            Begin VB.Label LblCompany 
               AutoSize        =   -1  'True
               Caption         =   "Pin Code :"
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
               Height          =   255
               Index           =   24
               Left            =   3960
               TabIndex        =   49
               Top             =   720
               Width           =   960
            End
         End
         Begin VB.Frame FraLegal 
            Height          =   2772
            Left            =   228
            TabIndex        =   44
            Top             =   528
            Width           =   8172
            Begin VB.ComboBox ComObjEdu 
               Height          =   360
               ItemData        =   "FClient.frx":233C
               Left            =   1440
               List            =   "FClient.frx":233E
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   2040
               Width           =   1335
            End
            Begin VB.ComboBox ComObjMed 
               Height          =   360
               ItemData        =   "FClient.frx":2340
               Left            =   4320
               List            =   "FClient.frx":2342
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   2040
               Width           =   1335
            End
            Begin VB.TextBox TxtTrustRegNo 
               Height          =   360
               Left            =   1560
               TabIndex        =   13
               Top             =   360
               Width           =   2895
            End
            Begin VB.TextBox TxtPan 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1320
               TabIndex        =   15
               Top             =   960
               Width           =   2655
            End
            Begin MSMask.MaskEdBox TxtTrustRegDt 
               Height          =   345
               Left            =   6120
               TabIndex        =   14
               ToolTipText     =   "Enter From Date"
               Top             =   360
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   609
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
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
               Caption         =   "Educational:"
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
               Index           =   2
               Left            =   120
               TabIndex        =   65
               Top             =   2040
               Width           =   1260
            End
            Begin VB.Label LblCompany 
               Caption         =   "Medical:"
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
               Index           =   23
               Left            =   3360
               TabIndex        =   64
               Top             =   2040
               Width           =   840
            End
            Begin VB.Label LblCompany 
               Caption         =   "Primary Objects:"
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
               Left            =   120
               TabIndex        =   63
               Top             =   1560
               Width           =   1575
            End
            Begin VB.Shape ShpMst 
               Height          =   2652
               Index           =   3
               Left            =   0
               Top             =   120
               Width           =   8172
            End
            Begin VB.Label LblCompany 
               Caption         =   "Trust Reg No.:"
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
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   47
               Top             =   360
               Width           =   1350
            End
            Begin VB.Label LblCompany 
               Caption         =   "Trust Reg Dt.:"
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
               Height          =   255
               Index           =   12
               Left            =   4680
               TabIndex        =   46
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label LblCompany 
               Caption         =   "Trust PAN:"
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
               Left            =   120
               TabIndex        =   45
               Top             =   1080
               Width           =   1035
            End
         End
         Begin VB.Frame FraBank 
            Height          =   1692
            Left            =   -74772
            TabIndex        =   40
            Top             =   528
            Width           =   6492
            Begin VB.TextBox TxtPrimaryBank 
               Height          =   360
               Left            =   2100
               TabIndex        =   18
               ToolTipText     =   "Enter Phone No."
               Top             =   240
               Width           =   3852
            End
            Begin VB.TextBox TxtPBankBranch 
               Height          =   360
               Left            =   960
               TabIndex        =   19
               Top             =   720
               Width           =   1692
            End
            Begin VB.TextBox TxtPBankAdd 
               Height          =   360
               Left            =   960
               TabIndex        =   20
               Top             =   1200
               Width           =   5472
            End
            Begin VB.Shape ShpMst 
               Height          =   1572
               Index           =   4
               Left            =   0
               Top             =   120
               Width           =   6492
            End
            Begin VB.Label LblCompany 
               Caption         =   "Primary Bank Name :"
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
               Index           =   3
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   1932
            End
            Begin VB.Label LblCompany 
               Caption         =   "Branch :"
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
               Index           =   1
               Left            =   120
               TabIndex        =   42
               Top             =   720
               Width           =   768
            End
            Begin VB.Label LblCompany 
               Caption         =   "Address :"
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
               Index           =   0
               Left            =   120
               TabIndex        =   41
               Top             =   1236
               Width           =   828
            End
         End
         Begin VB.Frame FraFcra 
            Height          =   2172
            Left            =   -74772
            TabIndex        =   34
            Top             =   528
            Width           =   7812
            Begin VB.TextBox TxtFCRSRegNo 
               Height          =   360
               Left            =   1560
               TabIndex        =   21
               Top             =   240
               Width           =   2772
            End
            Begin VB.TextBox TxtFCRABank 
               Height          =   360
               Left            =   1980
               TabIndex        =   23
               Top             =   720
               Width           =   3852
            End
            Begin VB.TextBox TxtFCRAType 
               Height          =   360
               Left            =   1680
               TabIndex        =   24
               Top             =   1200
               Width           =   1572
            End
            Begin VB.TextBox TxtFCRAAcNo 
               Height          =   360
               Left            =   1500
               TabIndex        =   25
               ToolTipText     =   "Enter Phone No."
               Top             =   1680
               Width           =   2652
            End
            Begin MSMask.MaskEdBox TxtFCRSRegDt 
               Height          =   345
               Left            =   6000
               TabIndex        =   22
               ToolTipText     =   "Enter From Date"
               Top             =   240
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   609
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##-##-####"
               PromptChar      =   " "
            End
            Begin VB.Shape ShpMst 
               Height          =   2052
               Index           =   5
               Left            =   0
               Top             =   120
               Width           =   7812
            End
            Begin VB.Label LblCompany 
               Caption         =   "FCRA Reg.No. :"
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
               Index           =   17
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   1416
            End
            Begin VB.Label LblCompany 
               Caption         =   "FCRA Reg.Dt.:"
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
               Index           =   18
               Left            =   4440
               TabIndex        =   38
               Top             =   240
               Width           =   1572
            End
            Begin VB.Label LblCompany 
               Caption         =   "FCRA Bank Name :"
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
               Index           =   19
               Left            =   120
               TabIndex        =   37
               Top             =   720
               Width           =   1728
            End
            Begin VB.Label LblCompany 
               Caption         =   "FCRA A/c.Type :"
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
               Index           =   20
               Left            =   120
               TabIndex        =   36
               Top             =   1200
               Width           =   1476
            End
            Begin VB.Label LblCompany 
               Caption         =   "FCRA A/c.No.:"
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
               Index           =   21
               Left            =   120
               TabIndex        =   35
               Top             =   1680
               Width           =   1308
            End
         End
      End
      Begin VB.Label LblCompany 
         Caption         =   "Press F4 Key For Parent Code Help"
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
         Left            =   360
         TabIndex        =   62
         Top             =   4320
         Width           =   5748
      End
      Begin VB.Shape ShpMst 
         Height          =   5412
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   15492
      End
   End
End
Attribute VB_Name = "FrmClientMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGroup As New ADODB.Recordset
Dim mActivity As String
Dim mAuto As Double

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then KeyAscii = 0
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub Form_Load()
    ClearText
    SetGrid
    If VsfHelp.Rows > 1 Then Display
    SetTool True
    Me.Left = 50
    Me.Top = 50
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mActivity = "Edit" Then
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
Private Function Display()
With VsfHelp
    If .Rows > 1 Then
        mAuto = Val(.TextMatrix(.Row, 0))
        TxtFileNo.Text = .TextMatrix(.Row, 2)
        TxtName.Text = .TextMatrix(.Row, 3)
        ComGroup.Text = .TextMatrix(.Row, 25)
        TxtParentCd.Text = .TextMatrix(.Row, 5)
        TxtTrustRegNo.Text = .TextMatrix(.Row, 6)
        If Len(Trim(.TextMatrix(.Row, 7))) = 10 Then TxtTrustRegDt.Text = .TextMatrix(.Row, 7)
        TxtPan.Text = .TextMatrix(.Row, 8)
        ComObjEdu.Text = IIf(.TextMatrix(.Row, 9) = -1, "Yes", "No")
        ComObjMed.Text = IIf(.TextMatrix(.Row, 10) = -1, "Yes", "No")
        TxtAdd.Text = .TextMatrix(.Row, 11)
        TxtCity.Text = .TextMatrix(.Row, 12)
        TxtTaluka.Text = .TextMatrix(.Row, 13)
        TxtDistrict.Text = .TextMatrix(.Row, 14)
        TxtPinCode.Text = .TextMatrix(.Row, 15)
        TxtState.Text = .TextMatrix(.Row, 16)
        TxtPrimaryBank.Text = .TextMatrix(.Row, 17)
        TxtPBankBranch.Text = .TextMatrix(.Row, 18)
        TxtPBankAdd.Text = .TextMatrix(.Row, 19)
        TxtFCRSRegNo.Text = .TextMatrix(.Row, 20)
        If Len(Trim(.TextMatrix(.Row, 21))) = 10 Then TxtFCRSRegDt.Text = .TextMatrix(.Row, 21)
        TxtFCRABank.Text = .TextMatrix(.Row, 22)
        TxtFCRAType.Text = .TextMatrix(.Row, 23)
        TxtFCRAAcNo.Text = .TextMatrix(.Row, 24)
        If Val(TxtParentCd.Text) <> 0 Then
            Dim RsQ As New ADODB.Recordset
            RsQ.Open "Select * From AcMst Where AcCode=" & Val(TxtParentCd.Text), DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsQ.EOF = False Then
                TxtParentName.Text = RsQ.Fields("AcName")
                TxtPFileNo.Text = RsQ.Fields("FileNo")
            Else
                TxtParentName.Text = ""
                TxtPFileNo.Text = ""
            End If
        Else
            TxtParentName.Text = ""
            TxtPFileNo.Text = ""
        End If
        SSTab1.Tab = 0
    End If
End With
End Function
Private Function SaveData()
On Error GoTo XErr
DbDataDB.BeginTrans
Select Case mActivity
    Case "Add"
        DbDataDB.Execute "Insert InTo AcMst (AcCode,AcName,Address,City,PinCode,State,Active,FileNo,AcType,TPan,TRegNo,TRegDt,Taluka,District,BBank," & _
        "BBranch,BAddress,FCRegNo,FCRegDt,FCBank,FCAcType,FCAcNo,PACode,ObjEdu,ObjMed) Values (" & mAuto & ",'" & TxtName.Text & "','" & TxtAdd.Text & _
        "','" & TxtCity.Text & "'," & Val(TxtPinCode.Text) & ",'" & TxtState.Text & "',-1,'" & TxtFileNo.Text & "'," & ComGroup.ItemData(ComGroup.ListIndex) & ",'" & _
        TxtPan.Text & "','" & TxtTrustRegNo.Text & "'," & IIf(Len(Trim(TxtTrustRegDt.Text)) = 10, "CDate('" & TxtTrustRegDt.Text & "')", "Null") & ",'" & TxtTaluka.Text & "','" & TxtDistrict.Text & "','" & TxtPrimaryBank.Text & _
        "','" & TxtPBankBranch.Text & "','" & TxtPBankAdd.Text & "','" & TxtFCRSRegNo.Text & "'," & IIf(Len(Trim(TxtFCRSRegDt.Text)) = 10, "CDate('" & TxtFCRSRegDt.Text & "')", "Null") & ",'" & TxtFCRABank.Text & "','" & _
        TxtFCRAType.Text & "','" & TxtFCRAAcNo.Text & "'," & Val(TxtParentCd.Text) & "," & IIf(ComObjEdu.Text = "Yes", -1, 0) & "," & IIf(ComObjMed.Text = "Yes", -1, 0) & ")"
        
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAuto & ",'CLIENT_MST','ADD_NEW','" & Date & "','" & Time & "')"
    Case "Edit"
        DbDataDB.Execute "Update AcMst Set AcName='" & TxtName.Text & "',Address='" & TxtAdd.Text & "',City='" & TxtCity.Text & "',PinCode=" & Val(TxtPinCode.Text) & _
        ",State='" & TxtState.Text & "',Active=-1,FileNo='" & TxtFileNo.Text & "',AcType=" & ComGroup.ItemData(ComGroup.ListIndex) & ",TPan='" & TxtPan.Text & "',TRegNo='" & _
        TxtTrustRegNo.Text & "',TRegDt=" & IIf(Len(Trim(TxtTrustRegDt.Text)) = 10, "CDate('" & TxtTrustRegDt.Text & "')", "Null") & ",Taluka='" & TxtTaluka.Text & _
        "',District='" & TxtDistrict.Text & "',BBank='" & TxtPrimaryBank.Text & "',BBranch='" & TxtPBankBranch.Text & "',BAddress='" & TxtPBankAdd.Text & "',FCRegNo='" & _
        TxtFCRSRegNo.Text & "',FCRegDt=" & IIf(Len(Trim(TxtFCRSRegDt.Text)) = 10, "CDate('" & TxtFCRSRegDt.Text & "')", "Null") & ",FCBank='" & TxtFCRABank.Text & "',FCAcType='" & _
        TxtFCRAType.Text & "',FCAcNo='" & TxtFCRAAcNo.Text & "',PACode=" & Val(TxtParentCd.Text) & ",ObjEdu=" & IIf(ComObjEdu.Text = "Yes", -1, 0) & _
        ",ObjMed=" & IIf(ComObjMed.Text = "Yes", -1, 0) & " Where AcCode=" & mAuto
    
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAuto & ",'CLIENT_MST','UPDATE','" & Date & "','" & Time & "')"
    Case "Delete"
        DbDataDB.Execute "Update AcMst Set Active=0, PACode=0 Where AcCode=" & mAuto
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAuto & ",'CLIENT_MST','" & ("DELETE_" & TxtFileNo.Text & "_" & TxtName.Text) & "','" & Date & "','" & Time & "')"
End Select
DbDataDB.CommitTrans
Exit Function
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Function
Private Function ClearText()
Dim ObjText As Object
For Each ObjText In Me
    If TypeOf ObjText Is TextBox Then ObjText.Text = ""
    TxtFCRSRegDt = "  -  -    "
    TxtTrustRegDt = "  -  -    "
Next
TxtState.Text = "Gujarat"
ComObjEdu.Clear
ComObjEdu.AddItem "No"
ComObjEdu.AddItem "Yes"
ComObjEdu.ListIndex = 0
ComObjMed.Clear
ComObjMed.AddItem "No"
ComObjMed.AddItem "Yes"
ComObjMed.ListIndex = 0
VsfAcHelp.Left = 18000
If ComGroup.ListCount > 0 Then ComGroup.ListIndex = 0
End Function
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Add"
        mActivity = "Add"
        SetTool False
        ClearText
        VsfHelp.Enabled = False
        SSTab1.Tab = 0
        TxtFileNo.SetFocus
    Case "Edit"
        If VsfHelp.Rows > 1 Then
            mActivity = "Edit"
            VsfHelp.Enabled = False
            SetTool False
            TxtFileNo.SetFocus
        Else
            MsgBox "Sorry! Record Not Found.", vbInformation, "Alert"
            VsfHelp.SetFocus
        End If
    Case "Delete"
        If VsfHelp.Rows > 1 Then
            mActivity = "Delete"
            VsfHelp.Enabled = False
            SetTool False
            TxtName.SetFocus
        Else
            MsgBox "Sorry! Record Not Found.", vbInformation, "Alert"
            VsfHelp.SetFocus
        End If
    Case "Save"
        If TxtName.Text = "" Then
            MsgBox "Pl.Enter Name.", vbInformation, "Alert"
            TxtName.SetFocus
        ElseIf ComGroup.Text = "" Then
            MsgBox "Pl.Group Name.", vbInformation, "Alert"
            ComGroup.SetFocus
        Else
            If mActivity = "Add" Then GAuto
            If mActivity <> "Delete" Then
                If CheckData = False Then SaveData Else Exit Sub
            Else
                Dim RsQData As New ADODB.Recordset
                RsQData.Open "Select AcCode from QGroup Where AcCode<>" & mAuto & " And (SACode=" & mAuto & " Or PACode=" & mAuto & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                If RsQData.EOF = False Then
                    MsgBox "Client has child branches in the software. Please delete child branches before deleting client.", vbCritical, "Alert"
                    mActivity = ""
                    ClearText
                    SetTool True
                    SetGrid
                    Display
                    VsfHelp.Enabled = True
                    VsfHelp.SetFocus
                    Exit Sub
                End If
                Set RsQData = Nothing
                RsQData.Open "Select AcCode From OpDtl Where AcCode=" & mAuto & " Union All Select AcCode From RpDtl Where AcCode=" & mAuto & " Union All Select AcCode From JvDtl " & _
                "Where AcCode=" & mAuto & " Union All Select AcCode From NtDtl Where AcCode=" & mAuto & " Union All Select AcCode From RepDtl Where AcCode=" & mAuto & _
                " Union All Select AcCode From 9CDtl Where AcCode=" & mAuto & " Union All Select AcCode From ArpDtl Where AcCode=" & mAuto, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                If RsQData.EOF = False Then
                    If MsgBox("Data found for this Client in the Software." + vbCrLf + "Do you want to Delete the Data and Client ?", vbCritical + vbYesNo, "Alert") = vbYes Then
                        DbDataDB.BeginTrans
                        DbDataDB.Execute "Delete From OpDtl Where AcCode=" & mAuto
                        DbDataDB.Execute "Delete From RpDtl Where AcCode=" & mAuto
                        DbDataDB.Execute "Delete From JvDtl Where AcCode=" & mAuto
                        DbDataDB.Execute "Delete From NtDtl Where AcCode=" & mAuto
                        DbDataDB.Execute "Delete From RepDtl Where AcCode=" & mAuto
                        DbDataDB.Execute "Delete From 9CDtl Where AcCode=" & mAuto
                        DbDataDB.Execute "Delete From EntMst Where AcCode=" & mAuto
                        DbDataDB.Execute "Delete From LedMst Where AcCode=" & mAuto
                        DbDataDB.Execute "Delete From ArpDtl Where AcCode=" & mAuto
                        DbDataDB.CommitTrans
                        SaveData
                    End If
                Else
                    SaveData
                End If
            End If
            mActivity = ""
            ClearText
            SetTool True
            SetGrid
            Display
            VsfHelp.Enabled = True
            VsfHelp.SetFocus
        End If
    Case "Cancel"
        mActivity = ""
        VsfHelp.Enabled = True
        SetTool True
        ClearText
        SetTool True
        SetGrid
        Display
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
If mUType <> "A" Then
    TlbSav(0).Enabled = False
    TlbSav(1).Enabled = False
    TlbSav(2).Enabled = False
End If
End Function
Private Function SetGrid()
Dim RsQ As New ADODB.Recordset
ComGroup.Clear
RsGroup.Open "Select * From GrpMst Order By GName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsGroup.EOF = False
    ComGroup.AddItem RsGroup.Fields("GName")
    ComGroup.ItemData(ComGroup.NewIndex) = RsGroup.Fields("GCode")
    RsGroup.MoveNext
Loop
ComGroup.ListIndex = 0
Set RsGroup = Nothing
RsQ.Open "Select AcMst.*,GrpMst.GName From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfHelp.DataSource = RsQ
With VsfHelp
    .FontSize = 11
    .ColWidth(0) = 0    'ACCODE
    .ColWidth(1) = 0    'ACTIVE
    .TextMatrix(0, 2) = "FILE NO."
    .ColWidth(2) = 1500
    .TextMatrix(0, 3) = "NAME"
    .ColWidth(3) = 4500
    .ColWidth(4) = 0    'TYPECODE
    .ColWidth(5) = 0   'PARENTCD
    .TextMatrix(0, 6) = "TRUST.REGNO"
    .ColWidth(6) = 1000
    .TextMatrix(0, 7) = "TRUST.REGDT"
    .ColWidth(7) = 1000
    .TextMatrix(0, 8) = "TRUST.PAN"
    .ColWidth(8) = 1000
    .ColWidth(9) = 0    'OBJEDU
    .ColWidth(10) = 0    'OBJMED
    .ColWidth(11) = 0   'ADDRESS
    .TextMatrix(0, 12) = "CITY"
    .ColWidth(12) = 1000
    .ColWidth(13) = 0   'TALUKA
    .TextMatrix(0, 14) = "DISTRICT"
    .ColWidth(14) = 1000
    .ColWidth(15) = 0   'PINCODE
    .ColWidth(16) = 0   'STATE
    .TextMatrix(0, 17) = "PRI.BANK"
    .ColWidth(17) = 1000
    .TextMatrix(0, 18) = "BRANCH"
    .ColWidth(18) = 1000
    .ColWidth(19) = 0   'PRIBANKADD
    .TextMatrix(0, 20) = "FCRAREGNO"
    .ColWidth(20) = 1000
    .TextMatrix(0, 21) = "FCRAREGDT"
    .ColWidth(21) = 1000
    .TextMatrix(0, 22) = "FCRAREGBANK"
    .ColWidth(22) = 1000
    .TextMatrix(0, 23) = "FCRAACTYPE"
    .ColWidth(23) = 1000
    .TextMatrix(0, 24) = "FCRAACCODE"
    .ColWidth(24) = 1000
    .Col = 2
    .ExtendLastCol = True
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
End With
Set RsQ = Nothing
RsQ.Open "Select AcMst.*,GrpMst.GName From AcMst,GrpMst Where AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfAcHelp.DataSource = RsQ
With VsfAcHelp
    .FontSize = 11
    .ColWidth(0) = 0    'ACCODE
    .ColWidth(1) = 0    'ACTIVE
    .TextMatrix(0, 2) = "FILE NO."
    .ColWidth(2) = 1500
    .TextMatrix(0, 3) = "NAME"
    .ColWidth(3) = 4500
    .ColWidth(4) = 0    'TYPECODE
    .ColWidth(5) = 0   'PARENTCD
    .TextMatrix(0, 6) = "TRUST.REGNO"
    .ColWidth(6) = 1000
    .TextMatrix(0, 7) = "TRUST.REGDT"
    .ColWidth(7) = 1000
    .TextMatrix(0, 8) = "TRUST.PAN"
    .ColWidth(8) = 1000
    .ColWidth(9) = 0    'OBJEDU
    .ColWidth(10) = 0    'OBJMED
    .ColWidth(11) = 0   'ADDRESS
    .TextMatrix(0, 12) = "CITY"
    .ColWidth(12) = 1000
    .ColWidth(13) = 0   'TALUKA
    .TextMatrix(0, 14) = "DISTRICT"
    .ColWidth(14) = 1000
    .ColWidth(15) = 0   'PINCODE
    .ColWidth(16) = 0   'STATE
    .TextMatrix(0, 17) = "PRI.BANK"
    .ColWidth(17) = 1000
    .TextMatrix(0, 18) = "BRANCH"
    .ColWidth(18) = 1000
    .ColWidth(19) = 0   'PRIBANKADD
    .TextMatrix(0, 20) = "FCRAREGNO"
    .ColWidth(20) = 1000
    .TextMatrix(0, 21) = "FCRAREGDT"
    .ColWidth(21) = 1000
    .TextMatrix(0, 22) = "FCRAREGBANK"
    .ColWidth(22) = 1000
    .TextMatrix(0, 23) = "FCRAACTYPE"
    .ColWidth(23) = 1000
    .TextMatrix(0, 24) = "FCRAACCODE"
    .ColWidth(24) = 1000
    .Col = 2
    .ExtendLastCol = True
    .Refresh
End With
End Function

Private Function GAuto()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcCode As RCode From AcMst Where Active=0", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQry.EOF = False Then
    mAuto = RsQry.Fields("RCode")
    mActivity = "Edit"
Else
    Set RsQry = Nothing
    RsQry.Open "Select IIF(IsNull(Max(AcCode))=True,1,Max(AcCode)+1) As RCode From AcMst", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    mAuto = RsQry.Fields("RCode")
End If
End Function

Private Sub TxtState_Validate(Cancel As Boolean)
    SSTab1.Tab = 2
    TxtTrustRegNo.SetFocus
End Sub

Private Sub TxtPan_Validate(Cancel As Boolean)
    SSTab1.Tab = 3
    TxtPrimaryBank.SetFocus
End Sub

Private Sub TxtPFileNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    If mActivity <> "" Then
        If VsfAcHelp.Rows > 1 Then
            VsfAcHelp.Left = 1800
            VsfAcHelp.SetFocus
        End If
    End If
End If
End Sub
Private Sub TxtParentCd_Validate(Cancel As Boolean)
    SSTab1.Tab = 1
    TxtAdd.SetFocus
End Sub
Private Sub TxtPBankAdd_Validate(Cancel As Boolean)
    SSTab1.Tab = 4
    TxtFCRSRegNo.SetFocus
End Sub
Private Sub VsfAcHelp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TxtParentCd.Text = VsfAcHelp.TextMatrix(VsfAcHelp.Row, 0)
        TxtParentName.Text = VsfAcHelp.TextMatrix(VsfAcHelp.Row, 3)
        TxtPFileNo.Text = VsfAcHelp.TextMatrix(VsfAcHelp.Row, 2)
        VsfAcHelp.Left = 18000
        TxtPFileNo.SetFocus
    ElseIf KeyCode = 27 Then
        VsfAcHelp.Left = 18000
        SSTab1.Tab = 1
        SSTab1.SetFocus
    End If
End Sub
Private Sub VsfHelp_RowColChange()
    If mActivity = "" And VsfHelp.Rows > 1 Then Display
End Sub

Private Function CheckData() As Boolean
Dim RsQ As New ADODB.Recordset
RsQ.Open "Select * From AcMst Where AcCode<>" & mAuto & " And FileNo='" & TxtFileNo.Text & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then
    MsgBox "Duplicate File No.", vbCritical, "Alert"
    CheckData = True
    Exit Function
End If
If TxtPan.Text <> "" Then
    Set RsQ = Nothing
    RsQ.Open "Select * From AcMst Where AcCode<>" & mAuto & " And TPan='" & TxtPan.Text & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.EOF = False Then
        MsgBox "Duplicate PAN", vbCritical, "Alert"
        CheckData = True
        Exit Function
    End If
End If
If TxtTrustRegNo.Text <> "" Then
    Set RsQ = Nothing
    RsQ.Open "Select * From AcMst Where AcCode<>" & mAuto & " And TRegNo='" & TxtTrustRegNo.Text & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.EOF = False Then
        MsgBox "Duplicate Trust Reg. No.", vbCritical, "Alert"
        CheckData = True
        Exit Function
    End If
End If
If TxtFCRSRegNo.Text <> "" Then
    Set RsQ = Nothing
    RsQ.Open "Select * From AcMst Where AcCode<>" & mAuto & " And FCRegNo='" & TxtFCRSRegNo.Text & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.EOF = False Then
        MsgBox "Duplicate FC Reg. No.", vbCritical, "Alert"
        CheckData = True
        Exit Function
    End If
End If
End Function
