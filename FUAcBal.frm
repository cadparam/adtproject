VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmUOpenBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening Balance"
   ClientHeight    =   9570
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   15690
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FUAcBal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   15690
   Begin VB.Frame FraItR 
      Height          =   7335
      Left            =   21000
      TabIndex        =   26
      Top             =   720
      Width           =   7215
      Begin VB.TextBox TxtITRTotal 
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   6840
         Width           =   1776
      End
      Begin VB.CommandButton CmdIClose 
         Caption         =   "C&lose"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Exit"
         Top             =   6840
         Width           =   950
      End
      Begin VB.CommandButton CmdISave 
         Caption         =   "S&ave"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Exit"
         Top             =   6840
         Width           =   950
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfItR 
         Height          =   6405
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   7000
         _cx             =   12347
         _cy             =   11298
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
         BackColor       =   16761087
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   128
         BackColorBkg    =   16777152
         BackColorAlternate=   16761087
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
         FormatString    =   $"FUAcBal.frx":000C
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
         Height          =   7215
         Index           =   4
         Left            =   0
         Top             =   120
         Width           =   7215
      End
   End
   Begin VB.Frame FraFurniture 
      Height          =   7335
      Left            =   21000
      TabIndex        =   25
      Top             =   840
      Width           =   7215
      Begin VB.TextBox TxtSFETotal 
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   6840
         Width           =   1776
      End
      Begin VB.CommandButton CmdFSave 
         Caption         =   "Sa&ve"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   6840
         Width           =   950
      End
      Begin VB.CommandButton CmdFClose 
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit"
         Top             =   6840
         Width           =   950
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfFurniture 
         Height          =   6405
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7000
         _cx             =   12347
         _cy             =   11298
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
         BackColor       =   16761087
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   128
         BackColorBkg    =   16777152
         BackColorAlternate=   16761087
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
         FormatString    =   $"FUAcBal.frx":0055
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
         Height          =   7215
         Index           =   3
         Left            =   0
         Top             =   120
         Width           =   7215
      End
   End
   Begin VB.Frame FraNote 
      Height          =   6612
      Left            =   18000
      TabIndex        =   14
      Top             =   0
      Width           =   10000
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
         Left            =   240
         Picture         =   "FUAcBal.frx":009E
         TabIndex        =   22
         ToolTipText     =   "Save"
         Top             =   5640
         Width           =   800
      End
      Begin VB.TextBox TxtOTotal 
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   5640
         Width           =   1600
      End
      Begin VB.TextBox TxtNTotal 
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
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5640
         Width           =   1600
      End
      Begin VB.TextBox TxtTotal 
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
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5640
         Width           =   1600
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Close"
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
         Left            =   3168
         Picture         =   "FUAcBal.frx":04E0
         TabIndex        =   3
         ToolTipText     =   "Cancel"
         Top             =   5640
         Width           =   975
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfNote 
         Height          =   5412
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9780
         _cx             =   17251
         _cy             =   9546
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
         BackColorBkg    =   12640511
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
         FormatString    =   $"FUAcBal.frx":0922
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
         Index           =   2
         Left            =   0
         Top             =   120
         Width           =   9972
      End
   End
   Begin VB.Frame FraClientHelp 
      Height          =   4212
      Left            =   18000
      TabIndex        =   12
      Top             =   1920
      Width           =   14750
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   3852
         Left            =   120
         TabIndex        =   13
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
         FormatString    =   $"FUAcBal.frx":096B
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
      Height          =   9500
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   15492
      Begin VB.CommandButton CmdITR 
         Caption         =   "&ITR"
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
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exit"
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
         Height          =   372
         Index           =   3
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdBF 
         Caption         =   "Balance Brought Forward"
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exit"
         Top             =   240
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export Data"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit"
         Top             =   8880
         Width           =   2064
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
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   9000
         Width           =   1776
      End
      Begin VB.TextBox TxtDTotal 
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   9000
         Width           =   1776
      End
      Begin VB.CommandButton CmdFurnityre 
         Caption         =   "School &Furniture Equipment"
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   2895
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
         Picture         =   "FUAcBal.frx":09B4
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
         TabIndex        =   9
         Top             =   240
         Width           =   4896
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   8200
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   15252
         _cx             =   26903
         _cy             =   14464
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
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FUAcBal.frx":0FEA
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
      Begin VB.Label LblTotal 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   6960
         TabIndex        =   21
         Top             =   9000
         Width           =   60
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
         TabIndex        =   11
         Top             =   240
         Width           =   1236
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   9375
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   15495
      End
   End
End
Attribute VB_Name = "FrmUOpenBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mAcCode As Double
Private Sub CmdBF_Click()
If MsgBox("Are you sure to brought-forward ledger balance from " & Format(CStr(Val(Mid(mYear, 9, 2)) - 1), "00") + "-" + Mid(mYear, 9, 2), vbInformation + vbYesNo) = vbYes Then
    Dim DBODataDB As New ADODB.Connection
    Dim RsOldData As New ADODB.Recordset
    Dim RsData As New ADODB.Recordset
    Dim mBal As Double
    DBODataDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & hPath + "1\" & Format(CStr(Val(Mid(mYear, 9, 2)) - 1), "00") + "-" + Mid(mYear, 9, 2) & "\DataDB.Mdb"
    DBODataDB.Open
    DbDataDB.BeginTrans
    DbDataDB.Execute "Delete From OpDtl"
    DbDataDB.CommitTrans
    DbDataDB.BeginTrans
    RsData.Open "Select AcCode,LCode,CBal+DBal As RSum From QTrialBal Where HType=1 And HCode Not In (5,57,58) Order By AcCode,LCode", DBODataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsData.EOF = False
        DbDataDB.Execute "Insert InTo OpDtl (AcCode,LCode,OpBal) Values (" & RsData.Fields("AcCode") & "," & RsData.Fields("LCode") & "," & RsData.Fields("RSum") & ")"
        RsData.MoveNext
    Loop
    DbDataDB.CommitTrans
    Set RsData = Nothing
    RsData.Open "Select AcCode,Sum(CBal+DBal) As RSum From QTrialBal Where HType=1 And HCode In (5,57,58) Group By AcCode Order By AcCode", DBODataDB, adOpenDynamic, adLockReadOnly, adCmdText
    DbDataDB.BeginTrans
    Do While RsData.EOF = False
        mBal = RsData.Fields("RSum") + SetProfit(RsData.Fields("AcCode"))
        DbDataDB.Execute "Insert InTo OpDtl (AcCode,LCode,OpBal) Values (" & RsData.Fields("AcCode") & ",1," & mBal & ")"
        RsData.MoveNext
    Loop
    DbDataDB.CommitTrans
    MsgBox "Successfully Record Updated.", vbInformation, "Alert"
End If
End Sub

Private Sub CmdFClose_Click()
    FraFurniture.Left = 21000
    CmdFurnityre.SetFocus
End Sub

Private Sub CmdIClose_Click()
    FraItR.Left = 21000
    CmdITR.SetFocus
End Sub
Private Sub CmdFSave_Click()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From OpDtl Where AcCode=" & mAcCode & " And LCode In (Select Distinct LCode From LedMst Where HCode=56)"
With VsfFurniture
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 Then
            DbDataDB.Execute "Insert InTo OpDtl (AcCode,LCode,OpBal) Values (" & mAcCode & "," & Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 1)) & ")"
        End If
    Next
End With
'DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'OPEN_SFE','UPDATE','" & Date & "','" & Time & "')"
DbDataDB.CommitTrans
MsgBox "Record Updated.", vbInformation, "Alert"
FraFurniture.Left = 21000
CmdFurnityre.SetFocus
Exit Sub
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Sub
Private Sub CmdISave_Click()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From OpDtl Where AcCode=" & mAcCode & " And LCode In (Select Distinct LCode From LedMst Where HCode=59)"
With VsfItR
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 Then
            DbDataDB.Execute "Insert InTo OpDtl (AcCode,LCode,OpBal) Values (" & mAcCode & "," & Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 1)) & ")"
        End If
    Next
End With
'DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'OPEN_ITR','UPDATE','" & Date & "','" & Time & "')"
DbDataDB.CommitTrans
MsgBox "Record Updated.", vbInformation, "Alert"
FraItR.Left = 21000
CmdITR.SetFocus
Exit Sub
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Sub

Private Sub CmdITR_Click()
If Val(LsvClient.TextMatrix(LsvClient.Row, 6)) = 0 And Val(LsvClient.TextMatrix(LsvClient.Row, 7)) = 1 Then
    FraItR.Left = 1200
    SetItR
    VsfItR.SetFocus
End If
End Sub
Private Sub CmdFurnityre_Click()
If Val(LsvClient.TextMatrix(LsvClient.Row, 7)) = 3 Then
    FraFurniture.Left = 1200
    SetFurniture
    VsfFurniture.SetFocus
End If
End Sub
Private Sub CmdSearch_Click()
    TlbSav(0).Enabled = True
    FraClientHelp.Left = 180
    LsvClient.SetFocus
End Sub
Private Sub Form_Load()
    Me.Left = 50
    Me.Top = 50
    SetCombo
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode,'' As RName,AcMst.AcType From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 And " & _
"AcMst.AcType=GrpMst.GCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.City,AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,''),GrpMst.GName,AcMst.AcCode," & _
"'',AcMst.AcType From AcMst,GrpMst,AcMst As AcMst1 Where AcMst.Active=-1 And AcMst.PaCode=AcMst1.AcCode And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    .ColWidth(5) = 0    'ACCODE
    .ColWidth(6) = 0    'PACODE
    .ColWidth(7) = 0    'ACTYPE
    .Col = 1
    .ExtendLastCol = True
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
End With
End Function
Private Function SetItR()
Dim RsQry As New ADODB.Recordset
Dim i As Integer
RsQry.Open "Select LName,0 As OpenBal,LCode From LedMst Where Active=-1 And HCode=59", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfItR.DataSource = RsQry
With VsfItR
    .TextMatrix(0, 0) = "PARTICULARS"
    .ColWidth(0) = 4800
    .TextMatrix(0, 1) = "OPENING BAL."
    .ColWidth(1) = 2000
    .ColFormat(1) = "0.00"
    .ColWidth(2) = 0    'LCODE
    .Col = 0
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
    Set RsQry = Nothing
    RsQry.Open "Select * From OpDtl Where AcCode=" & mAcCode & " And LCode In (Select Distinct LCode From LedMst Where HCode=59)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .Row = 1
        .Row = .FindRow(RsQry.Fields("LCode"), 1, 2)
        If .Row >= 1 Then .TextMatrix(.Row, 1) = RsQry.Fields("OpBal")
        RsQry.MoveNext
    Loop
    TxtITRTotal.Text = "0.00"
    For i = 1 To .Rows - 1
        TxtITRTotal = Format((Val(TxtITRTotal) + Val(.TextMatrix(i, 1))), "0.00")
    Next
End With
End Function
Private Function SetFurniture()
Dim RsQry As New ADODB.Recordset
Dim i As Integer
RsQry.Open "Select LName,0 As OpenBal,LCode From LedMst Where Active=-1 And HCode=56", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfFurniture.DataSource = RsQry
With VsfFurniture
    .TextMatrix(0, 0) = "NAME"
    .ColWidth(0) = 4800
    .TextMatrix(0, 1) = "OPENING BAL."
    .ColWidth(1) = 2000
    .ColFormat(1) = "0.00"
    .ColWidth(2) = 0    'LCODE
    .Col = 0
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
    Set RsQry = Nothing
    RsQry.Open "Select * From OpDtl Where AcCode=" & mAcCode & " And LCode In (Select Distinct LCode From LedMst Where HCode=56)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .Row = 1
        .Row = .FindRow(RsQry.Fields("LCode"), 1, 2)
        If .Row >= 1 Then .TextMatrix(.Row, 1) = RsQry.Fields("OpBal")
        RsQry.MoveNext
    Loop
    TxtSFETotal.Text = "0.00"
    For i = 1 To .Rows - 1
        TxtSFETotal = Format((Val(TxtSFETotal) + Val(.TextMatrix(i, 1))), "0.00")
    Next
End With
End Function

Private Sub ClearText()
    TxtName.Text = ""
    mAcCode = 0
    FraClientHelp.Left = 18000
    FraNote.Left = 18000
End Sub
Private Function SaveNote()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
With VsfNote
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Delete From OpDtl Where AcCode=" & mAcCode & " And LCode=" & Val(.TextMatrix(i, 4))
        If Val(VsfNote.TextMatrix(i, 1)) <> 0 Then
            DbDataDB.Execute "Insert InTo OpDtl (AcCode,LCode,OpBal) Values (" & mAcCode & "," & Val(.TextMatrix(i, 4)) & "," & Val(.TextMatrix(i, 1)) & ")"
        End If
    Next
End With
'DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'OPEN_BAL','UPDATE','" & Date & "','" & Time & "')"
DbDataDB.CommitTrans
Exit Function
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDIWork.PctMdi.Visible = True
    Unload Me
End Sub

Private Sub LsvClient_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtName.Text = LsvClient.TextMatrix(LsvClient.Row, 0)
    mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
    FraClientHelp.Left = 18000
    SetData
    CheckState
    VsfHelp.SetFocus
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'OPEN_BAL','VIEW_DATA','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
End If
End Sub
Private Sub VsfHelp_DblClick()
    SetLedger VsfHelp.Col
    FraNote.Left = 180
    VsfNote.SetFocus
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Save"
        If TxtName.Text = "" Then
            MsgBox "Sorry! Not Allowed.", vbInformation, "Black Data Error"
            TxtName.SetFocus
        Else
            If MsgBox("Are you sure to save?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
                SaveNote
                FraNote.Left = 18000
                SetData
                CmdSearch.SetFocus
            End If
        End If
    Case "Cancel"
        FraNote.Left = 18000
        VsfHelp.SetFocus
    Case "Exit"
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(0).Enabled = mVal
    TlbSav(1).Enabled = True
End Function
Private Sub SetData()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & _
mAcCode & ") And SeqMst.HCode=HedMst.HCode And HedMst.HType=1 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With VsfHelp
    .Cols = 8
    .Rows = 1
    .TextMatrix(0, 0) = "FUNDS AND LIABILITIES"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "NOTE"
    .ColWidth(1) = 700
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1700
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "PROPERTY AND ASSETS"
    .ColWidth(3) = 5000
    .TextMatrix(0, 4) = "NOTE"
    .ColWidth(4) = 700
    .TextMatrix(0, 5) = "AMOUNT RS"
    .ColWidth(5) = 1700
    .ColAlignment(5) = flexAlignRightCenter
    .TextMatrix(0, 6) = "DSIDE"
    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "CSIDE"
    .ColWidth(7) = 0
    .Refresh
    .Rows = 2
    .Row = 1
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 0) = RsQry.Fields("HName")
        .TextMatrix(.Row, 6) = RsQry.Fields("HCode")
        RsQry.MoveNext
        .Rows = .Rows + 1
        .Row = .Rows - 1
    Loop
    .Rows = .Rows + 1
    .Row = 1
    Set RsQry = Nothing
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & _
    mAcCode & ") And SeqMst.HCode=HedMst.HCode And HedMst.HType=1 And HedMst.HSide='D' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 3) = RsQry.Fields("HName")
        .TextMatrix(.Row, 7) = RsQry.Fields("HCode")
        RsQry.MoveNext
        If RsQry.EOF = False Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        End If
    Loop
    Dim i As Double
    Set RsQry = Nothing
    RsQry.Open "Select HedMst.HCode As HeadCode,Sum(OpDtl.OpBal) As TOpen From OpDtl,LedMst,HedMst Where OpDtl.AcCode=" & mAcCode & _
    " And OpDtl.LCode=LedMst.LCode And HedMst.HCode=LedMst.HCode Group By HedMst.HCode Order By HedMst.HCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
        If i > 0 Then
            .TextMatrix(i, 2) = Format(RsQry.Fields("TOpen"), "0.00")
        Else
            i = .FindRow(RsQry.Fields("HeadCode"), 1, 7)
            If i > 0 Then .TextMatrix(i, 5) = Format(RsQry.Fields("TOpen"), "0.00")
        End If
        RsQry.MoveNext
    Loop
End With
SetFinalTot
End Sub
Private Sub SetLedger(ByVal mSide As Integer)
Dim RsQ As New ADODB.Recordset
Dim RsLedger As New ADODB.Recordset
Dim mHeadCode As Double
If mSide <= 2 Then mHeadCode = Val(VsfHelp.TextMatrix(VsfHelp.Row, 6)) Else mHeadCode = Val(VsfHelp.TextMatrix(VsfHelp.Row, 7))
Set RsLedger = Nothing
RsLedger.Open "Select LedMst.LName,LedMst.LCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 " & _
"And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And LedMst.HCode=" & mHeadCode & _
" And LedMst.HCode=SeqMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & _
") And SeqMst.HCode=HedMst.HCode Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With VsfNote
    .Editable = flexEDNone
    .Cols = 5
    .Rows = 1
    .Row = 0
    .TextMatrix(0, 0) = "LEDGER"
    .ColWidth(0) = 4500
    .TextMatrix(0, 1) = "OPENING"
    .ColWidth(1) = 1600
    .ColFormat(1) = "0.00"
    .ColAlignment(1) = flexAlignRightCenter
    .FixedAlignment(1) = flexAlignRightCenter
    .TextMatrix(0, 2) = "CUR.YEAR"
    .ColWidth(2) = 1600
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .FixedAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "BALANCE"
    .ColWidth(3) = 1600
    .ColFormat(3) = "0.00"
    .ColAlignment(3) = flexAlignRightCenter
    .FixedAlignment(3) = flexAlignRightCenter
    .ColWidth(4) = 100  '  LEDGERCODE
    .Col = 0
    .Refresh
    .Rows = 2
    .Row = 1
    Set RsQ = Nothing
    RsQ.Open "Select * From OpDtl Where AcCode=" & mAcCode & " Order By LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsLedger.EOF = False
        .TextMatrix(.Row, 0) = RsLedger.Fields("LName")
        .TextMatrix(.Row, 4) = RsLedger.Fields("LCode")
        If RsQ.BOF = False Then
            RsQ.MoveFirst
            RsQ.Find "LCode=" & RsLedger.Fields("LCode")
        End If
        If RsQ.EOF = False Then .TextMatrix(.Row, 1) = RsQ.Fields("OpBal")
        RsLedger.MoveNext
        If RsLedger.EOF = False Then
            .Rows = .Rows + 1
            .Row = .Row + 1
        End If
    Loop
    SetTotal
End With
End Sub
Private Sub VsfFurniture_EnterCell()
With VsfFurniture
    If .Col = 0 Then
        .Editable = flexEDNone
        .AutoSearch = flexSearchFromCursor
    ElseIf .Col = 1 Then
        .AutoSearch = flexSearchNone
        .Editable = flexEDKbd
    End If
End With
End Sub

Private Sub VsfFurniture_RowColChange()
With VsfFurniture
    If .Col = 0 Then
        .Editable = flexEDNone
        .AutoSearch = flexSearchFromCursor
    ElseIf .Col = 1 Then
        .AutoSearch = flexSearchNone
        .Editable = flexEDKbd
    End If
End With
End Sub
Private Sub VsfFurniture_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Integer
TxtSFETotal.Text = "0.00"
With VsfFurniture
    For i = 1 To .Rows - 1
        TxtSFETotal = Format((Val(TxtSFETotal) + Val(.TextMatrix(i, 1))), "0.00")
    Next
End With
End Sub

Private Sub VsfItR_EnterCell()
With VsfItR
    If .Col = 0 Then
        .Editable = flexEDNone
        .AutoSearch = flexSearchFromCursor
    ElseIf .Col = 1 Then
        .AutoSearch = flexSearchNone
        .Editable = flexEDKbd
    End If
End With
End Sub

Private Sub VsfItR_RowColChange()
With VsfItR
    If .Col = 0 Then
        .Editable = flexEDNone
        .AutoSearch = flexSearchFromCursor
    ElseIf .Col = 1 Then
        .AutoSearch = flexSearchNone
        .Editable = flexEDKbd
        End If
End With
End Sub

Private Sub VsfItR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Integer
TxtITRTotal.Text = "0.00"
With VsfItR
    For i = 1 To .Rows - 1
        TxtITRTotal = Format((Val(TxtITRTotal) + Val(.TextMatrix(i, 1))), "0.00")
    Next
End With
End Sub

Private Sub VsfNote_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SetTotal
End Sub
Private Sub VsfNote_EnterCell()
If VsfNote.Col = 0 Or VsfNote.Col = 3 Then
    VsfNote.Editable = flexEDNone
    VsfNote.AutoSearch = flexSearchFromCursor
Else
    VsfNote.Editable = flexEDKbd
    VsfNote.AutoSearch = flexSearchNone
End If
End Sub
Private Sub VsfNote_RowColChange()
With VsfNote
    If .Col = 0 Or .Col = 3 Then
        .Editable = flexEDNone
        .AutoSearch = flexSearchFromCursor
    Else
        .Editable = flexEDKbd
        .AutoSearch = flexSearchNone
    End If
    If Mid(VsfHelp.TextMatrix(VsfHelp.Row, 3), 1, 4) = "Cash" Or Mid(VsfHelp.TextMatrix(VsfHelp.Row, 3), 1, 4) = "Bank" Then
    Else
        .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row, 1)) + Val(.TextMatrix(.Row, 2))
    End If
End With
End Sub
Private Sub VsfNote_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    SetTotal
End Sub
Private Sub SetTotal()
Dim i As Double
TxtOTotal.Text = "0.00"
TxtTotal.Text = "0.00"
TxtNTotal.Text = "0.00"
With VsfNote
    For i = 1 To .Rows - 1
        TxtOTotal.Text = Val(TxtOTotal.Text) + Val(.TextMatrix(i, 1))
        TxtTotal.Text = Val(TxtTotal.Text) + Val(.TextMatrix(i, 2))
        If Mid(VsfHelp.TextMatrix(VsfHelp.Row, 3), 1, 4) = "Cash" Or Mid(VsfHelp.TextMatrix(VsfHelp.Row, 3), 1, 4) = "Bank" Then
        
        Else
            .TextMatrix(i, 3) = Val(.TextMatrix(i, 1)) + Val(.TextMatrix(i, 2))
        End If
        TxtNTotal.Text = Val(TxtNTotal.Text) + Val(.TextMatrix(i, 3))
    Next
End With
TxtOTotal.Text = Format(TxtOTotal.Text, "0.00")
TxtTotal.Text = Format(TxtTotal.Text, "0.00")
TxtNTotal.Text = Format(TxtNTotal.Text, "0.00")
End Sub
Private Sub SetFinalTot()
Dim i As Double
TxtDTotal.Text = "0.00"
TxtCTotal.Text = "0.00"
With VsfHelp
    For i = 1 To .Rows - 1
        TxtDTotal.Text = Val(TxtDTotal.Text) + Val(.TextMatrix(i, 2))
        TxtCTotal.Text = Val(TxtCTotal.Text) + Val(.TextMatrix(i, 5))
    Next
End With
TxtDTotal.Text = Format(TxtDTotal.Text, "0.00")
TxtCTotal.Text = Format(TxtCTotal.Text, "0.00")
If Val(TxtCTotal.Text) - Val(TxtDTotal.Text) = 0 Then LblTotal.Caption = "" Else LblTotal.Caption = "Total mismatching Of Rs.-->" & Format(CStr(Abs(Val(TxtCTotal.Text) - Val(TxtDTotal.Text))), "0.00")
End Sub
Private Sub CheckState()
Dim RsQ As New ADODB.Recordset
Dim mString As String
RsQ.Open "Select * From QGroup Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
mString = CStr(RsQ.Fields("AcCode")) + "," + CStr(RsQ.Fields("PaCode")) + "," + CStr(RsQ.Fields("SaCode"))
Set RsQ = Nothing
RsQ.Open "Select * From RepDtl Where AcCode In (" & mString & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    If RsQ.Fields("RUDIN") <> "" Then
        MsgBox "UDIN generated. Data can not be edited.", vbCritical, "Alert"
        TlbSav(0).Enabled = False
        Exit Sub
    End If
    RsQ.MoveNext
Loop
End Sub
