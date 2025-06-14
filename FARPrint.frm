VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmAuditReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audit Report"
   ClientHeight    =   7575
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FARPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10860
   Begin VB.Frame FraNote 
      Height          =   6612
      Left            =   18000
      TabIndex        =   15
      Top             =   0
      Width           =   10000
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Picture         =   "FARPrint.frx":000C
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   5640
         Width           =   975
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfNote 
         Height          =   5412
         Left            =   120
         TabIndex        =   8
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
         FormatString    =   $"FARPrint.frx":044E
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
      TabIndex        =   13
      Top             =   1920
      Width           =   14750
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   3852
         Left            =   120
         TabIndex        =   14
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
         FormatString    =   $"FARPrint.frx":0497
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
      Height          =   7455
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   15495
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
         Left            =   6840
         Picture         =   "FARPrint.frx":04E0
         TabIndex        =   36
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   800
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
         Index           =   2
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   855
      End
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
         Height          =   372
         Index           =   3
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtOMetter 
         Height          =   1590
         Left            =   1680
         TabIndex        =   6
         Top             =   5400
         Width           =   8010
      End
      Begin VB.TextBox TxtEmphasis 
         Height          =   1470
         Left            =   1680
         TabIndex        =   5
         Top             =   3720
         Width           =   8010
      End
      Begin VB.TextBox TxtQualify 
         Height          =   1470
         Left            =   1680
         TabIndex        =   4
         Top             =   2040
         Width           =   8010
      End
      Begin VB.ComboBox ComQualify 
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   1440
      End
      Begin VB.ComboBox ComFAsset 
         Height          =   360
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2520
      End
      Begin VB.ComboBox ComAcMethod 
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2400
      End
      Begin VB.Frame FraMainExp 
         BackColor       =   &H00F7E0FE&
         Height          =   5412
         Left            =   1.80000e5
         TabIndex        =   25
         Top             =   480
         Width           =   15612
         Begin VB.CommandButton CmpMEClose 
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
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Cancel"
            Top             =   4920
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfMainExport 
            Height          =   5052
            Left            =   0
            TabIndex        =   27
            Top             =   240
            Width           =   15480
            _cx             =   27305
            _cy             =   8911
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
            BackColorBkg    =   15007437
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
            FormatString    =   $"FARPrint.frx":0922
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
            Height          =   5292
            Index           =   5
            Left            =   0
            Top             =   120
            Width           =   15612
         End
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
         Left            =   17280
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   6600
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
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   6600
         Width           =   1776
      End
      Begin VB.Frame FraDetail 
         Height          =   6132
         Left            =   18000
         TabIndex        =   19
         Top             =   0
         Width           =   8052
         Begin VB.TextBox TxtTotalD 
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
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   5640
            Width           =   1500
         End
         Begin VB.CommandButton CmdCancelD 
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
            Left            =   1248
            Picture         =   "FARPrint.frx":096B
            TabIndex        =   21
            ToolTipText     =   "Cancel"
            Top             =   5640
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfDetail 
            Height          =   5412
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   7860
            _cx             =   13864
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
            FormatString    =   $"FARPrint.frx":0DAD
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
            Height          =   6012
            Index           =   3
            Left            =   0
            Top             =   120
            Width           =   8052
         End
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
         Picture         =   "FARPrint.frx":0DF6
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
         TabIndex        =   10
         Top             =   240
         Width           =   4896
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   5895
         Left            =   18840
         TabIndex        =   7
         Top             =   720
         Width           =   15255
         _cx             =   26903
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
         FormatString    =   $"FARPrint.frx":142C
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
         TabIndex        =   33
         Top             =   5400
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
         TabIndex        =   32
         Top             =   3720
         Width           =   1275
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
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1485
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
         TabIndex        =   30
         Top             =   1440
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
         TabIndex        =   29
         Top             =   840
         Width           =   2325
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
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   2055
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
         TabIndex        =   12
         Top             =   240
         Width           =   1236
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   7335
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   10695
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
End
Attribute VB_Name = "FrmAuditReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim mAcCode As Double
Dim mAcList As String
Dim RsClient As New ADODB.Recordset
Dim RsRep As New ADODB.Recordset
Dim RsState As New ADODB.Recordset
Dim mProfit As Double
Dim mTrack As Boolean
Dim CState As String
Private Sub CmdSearch_Click()
    FraClientHelp.Left = 180
    LsvClient.SetFocus
End Sub

Private Sub ComQualify_Validate(Cancel As Boolean)
    If ComQualify.Text = "No" And TxtQualify.Text <> "" Then
        If MsgBox("Are you sure to remove qualification?", vbCritical + vbYesNo) = vbYes Then TxtQualify.Text = "" Else ComQualify.Text = "Yes"
    End If
End Sub

Private Sub Form_Load()
    Me.Left = 50
    Me.Top = 50
    DBWorkTmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & App.Path + "\LocalDB.Mdb'"
    DBWorkTmp.Open
    SetCombo
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 " & _
"And AcMst.AcType=1 And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    .TextMatrix(0, 5) = "ACCODE"
    .ColWidth(5) = 0
    .Col = 1
    .ExtendLastCol = True
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
End With
Set RsClient = Nothing
RsClient.Open "Select * From AcMst Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    TxtName.Text = ""
    mAcCode = 0
    FraClientHelp.Left = 18000
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDIWork.PctMdi.Visible = True
    Unload Me
End Sub
Private Sub LsvClient_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtName.Text = LsvClient.TextMatrix(LsvClient.Row, 0)
    mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
    If RsClient.BOF = False Then
        RsClient.MoveFirst
        RsClient.Find "AcCode=" & mAcCode
    End If
    FraClientHelp.Left = 18000
    Display
    SetParent
    Dim RsQ As New ADODB.Recordset
    DBWorkTmp.BeginTrans
    DBWorkTmp.Execute "Delete From TmpTrialBal"
    DBWorkTmp.Execute "Delete From TmpCtDtl"
    DBWorkTmp.CommitTrans
    DBWorkTmp.BeginTrans
    RsQ.Open "Select * From QTrialBal Where AcCode In (" & mAcList & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpTrialBal (HType,HSide,AcCode,HCode,LCode,OpDr,OpCr,ADr,ACr,DBal,CBal) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & _
        "'," & RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("OpDr") & "," & RsQ.Fields("OpCr") & "," & RsQ.Fields("ADr") & _
        "," & RsQ.Fields("ACr") & "," & RsQ.Fields("DBal") & "," & RsQ.Fields("CBal") & ")"
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From QCtDtl Where AcCode In (" & mAcList & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpCtDtl (HType,HSide,AcCode,HCode,LCode,ECode,EName,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
        RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("EName") & "','" & RsQ.Fields("Side") & _
        "'," & RsQ.Fields("Amt") & ")"
        RsQ.MoveNext
    Loop
    DBWorkTmp.CommitTrans
    mProfit = 0
    mTrack = False
    ComAcMethod.SetFocus
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Save"
        If mAcCode = 0 Then
            MsgBox "Please select client.", vbCritical, "Alert"
            Exit Sub
        End If
        SaveData
    Case "Print"
        If mAcCode = 0 Then
            MsgBox "Please select client.", vbCritical, "Alert"
            Exit Sub
        End If
        SaveData
        PrintRec
    Case "Exit"
        If MsgBox("Close All Reports? ", vbInformation + vbYesNo, "Confirmation") = vbYes Then
        DBWorkTmp.Close
        Unload Me
        End If
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(2).Enabled = True
End Function
Private Sub SetData()
Dim RsExclude As New ADODB.Recordset
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And SeqMst.HCode=" & _
"HedMst.HCode And HedMst.HType=1 And HedMst.HCode<>14 And HedMst.HSide='C' And HedMst.HCode<>14 Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And SeqMst.HCode=" & _
    "HedMst.HCode And HedMst.HType=1 And HedMst.HSide='D' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    RsQry.Open "Select HSide As AcSide,HCode As HeadCode,Sum(CBal) As TotRs From TmpTrialBal Where HType=1 And HCode<>14 Group By HSide,HCode Union All " & _
    "Select HSide,HCode,Sum(DBal) From TmpTrialBal Where HType=1 And HCode<>14 Group By HSide,HCode Order By HeadCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        If RsQry.Fields("AcSide") = "C" Then
            If mFinYear = "19-20" Then
                If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 57 Then
                    i = .FindRow(5, 1, 6)
                ElseIf RsQry.Fields("HeadCode") = 13 Then
                    i = .FindRow(1, 1, 6)
                Else
                    i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
                End If
            Else
                If RsQry.Fields("HeadCode") = 13 Then i = .FindRow(1, 1, 6) Else i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
            End If
            If i > 0 Then
                If mFinYear = "19-20" Then
                    If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 13 Or RsQry.Fields("HeadCode") = 5 Then    '   Income And Expenditure
                        If mProfit = 0 Then
                            mProfit = SetProfitAll(mAcList)
                            .TextMatrix(i, 2) = CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs") + mProfit)
                        Else
                            .TextMatrix(i, 2) = CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs"))
                        End If
                        .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2))), "0.00")
                    Else
                        .TextMatrix(i, 2) = CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs"))
                        .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "0.00")
                    End If
                Else
                    If RsQry.Fields("HeadCode") = 5 Then    '   Income And Expenditure
                        If mProfit = 0 Then
                            mProfit = SetProfitAll(mAcList)
                            .TextMatrix(i, 2) = CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs") + mProfit)
                        Else
                            .TextMatrix(i, 2) = CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs"))
                        End If
                        .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2))), "0.00")
                    Else
                        .TextMatrix(i, 2) = CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs"))
                        .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "0.00")
                    End If
                End If
            End If
        Else
            i = .FindRow(RsQry.Fields("HeadCode"), 1, 7)
            If i > 0 Then
                .TextMatrix(i, 5) = Format(CStr(Val(.TextMatrix(i, 5)) + RsQry.Fields("TotRs")), "0.00")
            End If
        End If
        RsQry.MoveNext
    Loop
    i = .FindRow(5, 1, 6)
    If i > 0 Then
        If Val(.TextMatrix(i, 2)) = 0 Then
            mProfit = SetProfitAll(mAcList)
            .TextMatrix(i, 2) = mProfit
        End If
    End If
    Set RsQry = Nothing
    RsQry.Open "Select HCode As HeadCode,Sum(CBal) As TotRs From TmpTrialBal Where LCode In(87,88,86,1) Group By HCode Order By HCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 57 Then
            i = .FindRow(5, 1, 6)
        Else
            i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
        End If
        If i > 0 Then
            If Val(.TextMatrix(i, 2)) = 0 Then
                If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 57 Or RsQry.Fields("HeadCode") = 5 Then     '   Income And Expenditure
                    mProfit = SetProfitAll(mAcList)
                    .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs")) + mProfit, "0.00")
                End If
            End If
        End If
        RsQry.MoveNext
    Loop
    Set RsQry = Nothing
    RsQry.Open "Select * From NtDtl Where AcCode=" & mAcCode & " And RType=1 Order By HCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        If RsQry.Fields("RSide") = "C" Then
            i = .FindRow(RsQry.Fields("HCode"), 1, 6)
            If i > 0 Then .TextMatrix(i, 1) = RsQry.Fields("Note")
        Else
            i = .FindRow(RsQry.Fields("HCode"), 1, 7)
            If i > 0 Then .TextMatrix(i, 4) = RsQry.Fields("Note")
        End If
        RsQry.MoveNext
    Loop
End With
SetFinalTot
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
End Sub
Private Sub PrintRec()
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
Set RsState = Nothing
RsState.Open "Select * from SActMst Where State='" & RsClient.Fields("State") & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpBSPrn"
DbDataDB.Execute "Delete From TmpNotePrn"
DbDataDB.Execute "Delete From TmpPLPrn"
DbDataDB.CommitTrans
SetData
BSheetPrint
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Balance Sheet does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
NotePrintBS
SetData1
PlAcPrint
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Income & Expenditure Account does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
NotePrintIE
If RsClient.BOF = False Then
    RsClient.MoveFirst
    RsClient.Find "AcCode=" & Val(LsvClient.TextMatrix(LsvClient.Row, 5))
End If
PrintData
If RsClient.Fields("State") = "Gujarat" Then Schedule9C
DbDataDB.BeginTrans
'DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'AUDIT_RPT','PRINT','" & Date & "','" & Time & "')"
DbDataDB.CommitTrans
ClearAll
End Sub
Private Sub SetParent()
Dim RsQ As New ADODB.Recordset
RsQ.Open "Select * From QGroup Where SaCode=" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
'mAcList = CStr(mAcCode)
mAcList = CStr(RsQ.Fields("AcCode"))
RsQ.MoveNext
Do While RsQ.EOF = False
    mAcList = mAcList + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
End Sub
Private Sub NotePrintBS()
Dim i As Integer
Dim mAcCodeAll As String
Dim mAcName As String
Dim mGroup As Double
Dim mTotal As Double
Dim RsDataO As New ADODB.Recordset
Dim RsDataD As New ADODB.Recordset
Dim RsQData As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim mRAmt As Double
Dim mSAmt As Double
Dim mCAmt As Double
Dim mDAmt As Double
DbDataDB.BeginTrans
mAcCodeAll = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName='Trust (Local Fund) Account'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName='Foreign Contribution Account'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcType=5 And AcName Not In ('Trust (Local Fund) Account','Foreign Contribution Account')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcType<>5 And AcName Not In ('" & LsvClient.TextMatrix(LsvClient.Row, 0) & _
"','Trust (Local Fund) Account','Foreign Contribution Account')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
Dim RsAc As New ADODB.Recordset
RsAc.Open "Select * From AcMst Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With VsfHelp
    Do While Len(mAcList) > 0
        If InStr(1, mAcList, ",") - 1 > 0 Then
            mAcCode = Val(Mid(mAcList, 1, InStr(1, mAcList, ",") - 1))
            mAcList = Mid(mAcList, InStr(1, mAcList, ",") + 1)
        Else
            mAcCode = Val(mAcList)
            mAcList = ""
        End If
        If RsAc.BOF = False Then
            RsAc.MoveFirst
            RsAc.Find "AcCode=" & mAcCode
        End If
        If RsAc.EOF = False Then
            If RsAc.Fields("AcName") = LsvClient.TextMatrix(LsvClient.Row, 0) Then
                mAcName = Space(3) + LsvClient.TextMatrix(LsvClient.Row, 0) + IIf(Len(RsAc.Fields("City")) <> 0, (", " + RsAc.Fields("City")), "")
            ElseIf RsAc.Fields("AcName") = "Trust Account" Then
                mAcName = Space(2) + RsAc.Fields("AcName")
            ElseIf RsAc.Fields("AcName") = "Foreign Contribution Account" Then
                mAcName = Space(1) + RsAc.Fields("AcName")
            Else
                mAcName = RsAc.Fields("AcName") + IIf(Len(RsAc.Fields("City")) <> 0, (", " + RsAc.Fields("City")), "")
            End If
        End If
        If Val(.TextMatrix(2, 1)) <> 0 Then
            i = .FindRow(2, 1, 6)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 6))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & _
                        "','" & RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(3, 1)) <> 0 Then
            i = .FindRow(3, 1, 6)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 6))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & _
                        "','" & RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(4, 1)) <> 0 Then
            i = .FindRow(4, 1, 6)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 6))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(1, 4)) <> 0 Then
            i = .FindRow(6, 1, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.OpDr As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And LedMst.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mRAmt = 0
                    mRAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From QTmpLike Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mSAmt = 0
                    mSAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mDAmt = 0
                    mDAmt = RsQData.Fields("RTotal") - mSAmt
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mCAmt = 0
                    mCAmt = RsQData.Fields("RTotal")
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,LName,SAmt,CAmt,DAmt,HCode,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & _
                    .TextMatrix(i, 3) & "'," & RsDataO.Fields("RTotal") & "," & mRAmt & ",'" & RsDataO.Fields("LName") & "'," & mSAmt & "," & mCAmt & "," & mDAmt & "," & mGroup & ",'" & mAcName & "')"
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(2, 4)) <> 0 Then
            i = .FindRow(7, 1, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(3, 4)) <> 0 Then
            i = .FindRow(8, 1, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.OpDr As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And LedMst.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & _
                    RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mRAmt = 0
                    mRAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From QTmpLike Where AcCode=" & mAcCode & _
                    " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mSAmt = 0
                    mSAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & _
                    RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mDAmt = 0
                    mDAmt = RsQData.Fields("RTotal") - mSAmt
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & _
                    RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mCAmt = 0
                    mCAmt = RsQData.Fields("RTotal")
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,LName,SAmt,CAmt,DAmt,HCode,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & _
                    .TextMatrix(i, 3) & "'," & RsDataO.Fields("RTotal") & "," & mRAmt & ",'" & RsDataO.Fields("LName") & "'," & mSAmt & "," & mCAmt & "," & mDAmt & "," & mGroup & ",'" & mAcName & "')"
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(4, 4)) <> 0 Then
            i = .FindRow(9, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,LedMst.HCode,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & _
                mAcCode & " And LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode<>14 And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & _
                        .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(5, 4)) <> 0 Then
            i = .FindRow(10, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                        "','" & RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(6, 4)) <> 0 Then
            i = .FindRow(11, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(7, 4)) <> 0 Then
            i = .FindRow(12, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And TmpTrialBal.LCode=LedMst.LCode And " & _
                "TmpTrialBal.LCode In (10000,10001,10315) Order By TmpTrialBal.LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("DBal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                        "',' " & RsDataO.Fields("LName") & "'," & RsDataO.Fields("DBal") & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
                Set RsDataO = Nothing
                RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And TmpTrialBal.HCode=" & mGroup & _
                " And TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.LCode Not In (10000,10001,10315) Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("DBal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                        "','" & RsDataO.Fields("LName") & "'," & RsDataO.Fields("DBal") & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
    If Val(.TextMatrix(1, 1)) <> 0 Then
        If mFinYear <> "19-20" Then
            i = .FindRow(1, 1, 6)
            If i = -1 Then i = .FindRow(13, 1, 6)
        Else
            i = .FindRow(57, 1, 6)
            If i = -1 Then i = .FindRow(1, 1, 6)
        End If
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            If Val(.TextMatrix(i, 6)) <> 1 Then
                RsDataO.Open "Select LedMst.LName,IIF(IsNull(OpCr)=True,0,OpCr) As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select LedMst.LName,IIF(IsNull(OpCr)=True,0,OpCr) As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode In (1,13)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            If RsDataO.EOF = False Then
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,LName,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','  Opening Balance'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & RsDataO.Fields("LName") & "','" & mAcName & "')"
                End If
            End If
            Set RsDataO = Nothing
            If Val(.TextMatrix(i, 6)) <> 1 Then
                RsDataO.Open "Select LedMst.LName,IIF(TmpCtDtl.Side=TmpCtDtl.HSide,'Add','Less') As TrnType,TmpCtDtl.EName,Sum(TmpCtDtl.Amt) As Amount From TmpCtDtl,LedMst Where TmpCtDtl.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpCtDtl.LCode And TmpCtDtl.HCode=" & mGroup & " Group By IIF(Side=HSide,'Add','Less'),LName,EName Order By IIF(Side=HSide,'Add','Less')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select LedMst.LName,IIF(TmpCtDtl.Side=TmpCtDtl.HSide,'Add','Less') As TrnType,TmpCtDtl.EName,Sum(TmpCtDtl.Amt) As Amount From TmpCtDtl,LedMst Where TmpCtDtl.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpCtDtl.LCode And TmpCtDtl.HCode In (1,13) Group By IIF(Side=HSide,'Add','Less'),LName,EName Order By IIF(Side=HSide,'Add','Less')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,LName,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' " & _
                    RsDataO.Fields("TrnType") & ":" & Space(1) & RsDataO.Fields("EName") & "'," & IIf(RsDataO.Fields("TrnType") = "Add", RsDataO.Fields("Amount"), RsDataO.Fields("Amount") * -1) & _
                    "," & IIf(RsDataO.Fields("TrnType") = "Add", 0, 1) & ",'" & RsDataO.Fields("LName") & "','" & mAcName & "')"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            If Val(.TextMatrix(i, 6)) <> 1 Then
                RsDataO.Open "Select LedMst.LName,IIF(IsNull(CBal)=True,0,CBal) As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select LedMst.LName,IIF(IsNull(CBal)=True,0,CBal) As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode In (1,13)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            If RsDataO.EOF = False Then
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,RName,LName,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'Closing Balance','" & RsDataO.Fields("LName") & "','" & mAcName & "')"
                End If
            End If
        End If
    End If
    Loop
    If Val(.TextMatrix(5, 1)) <> 0 Then
        i = .FindRow(5, 1, 6)
        If i = -1 Then i = .FindRow(58, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            If mFinYear = "19-20" Then
                RsDataO.Open "Select IIF(IsNull(Sum(OpCr))=True,0,Sum(OpCr)) As RTotal From TmpTrialBal Where HCode In (58,57,13,5)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select IIF(IsNull(Sum(OpCr))=True,0,Sum(OpCr)) As RTotal From TmpTrialBal Where HCode In (58,57,5)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            If RsDataO.EOF = False Then
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','  Opening Balance'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
            End If
            Set RsDataO = Nothing
            If mFinYear = "19-20" Then
                RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName, Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where HCode In (5,13,57,58) " & _
                "And ECode Not In (960,961,962,964) Group By IIF(Side=HSide,'Add','Less'),EName Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName, Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where HCode In (5,57,58) " & _
                "Group By IIF(Side=HSide,'Add','Less'),EName Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' " & _
                    RsDataO.Fields("TrnType") & ":" & Space(1) & RsDataO.Fields("EName") & "'," & RsDataO.Fields("Amount") & "," & IIf(RsDataO.Fields("TrnType") = "Add", 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            mTotal = SetProfitAll(mAcCodeAll)
            DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
            IIf(Round(mTotal, 2) >= 0, "Add: Surplus brought from Income and Expenditure Account", "Less: Deficit brought from Income and Expenditure Account") & _
            "'," & mTotal & "," & IIf(Round(mTotal, 2) >= 0, 0, 1) & ",'')"
        End If
    End If
End With
DbDataDB.CommitTrans
mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
SetParent
End Sub
Private Sub BSheetPrint()
Dim i As Double
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpBSPrn"
With VsfHelp
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Insert Into TmpBSPrn (LName,LAmt,LNote,RName,RAmt,RNote,SrN) Values ('" & .TextMatrix(i, 0) & "'," & _
        Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 3) & "'," & Val(.TextMatrix(i, 5)) & "," & _
        Val(.TextMatrix(i, 4)) & "," & i & ")"
    Next
End With
DbDataDB.CommitTrans
SetParent
End Sub
Private Sub PlAcPrint()
Dim i As Double
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpPLPrn"
DbDataDB.CommitTrans
DbDataDB.BeginTrans
With VsfHelp
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Insert Into TmpPLPrn (LName,LAmt,LNote,RName,RAmt,RNote,SrN) Values ('" & IIf(Len(.TextMatrix(i, 0)) = 0, "", IIf(Mid(.TextMatrix(i, 0), 1, 1) = "(", "", "To ") + .TextMatrix(i, 0)) & "'," & _
        Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 1)) & ",'" & IIf(Len(.TextMatrix(i, 3)) = 0, "", "By " + .TextMatrix(i, 3)) & "'," & Val(.TextMatrix(i, 5)) & "," & _
        Val(.TextMatrix(i, 4)) & "," & i & ")"
    Next
End With
DbDataDB.CommitTrans
If RsClient.BOF = False Then
    RsClient.MoveFirst
    RsClient.Find "AcCode=" & Val(LsvClient.TextMatrix(LsvClient.Row, 5))
End If
SetParent
End Sub
Private Sub SetData1()
Dim RsQry As New ADODB.Recordset
Dim i As Integer
With VsfHelp
    .Cols = 8
    .Rows = 1
    .TextMatrix(0, 0) = "EXPENDITURE"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "NOTE"
    .ColWidth(1) = 700
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1700
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "INCOME"
    .ColWidth(3) = 5000
    .TextMatrix(0, 4) = "NOTE"
    .ColWidth(4) = 700
    .TextMatrix(0, 5) = "AMOUNT RS"
    .ColWidth(5) = 1700
    .ColAlignment(5) = flexAlignRightCenter
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .Refresh
    .Rows = 2
    .Row = 1
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And SeqMst.HCode=HedMst.HCode And HedMst.HType=0 And HedMst.HSide='D' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 0) = RsQry.Fields("HName")
        .TextMatrix(.Row, 6) = RsQry.Fields("HCode")
        RsQry.MoveNext
        If RsQry.EOF = False Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
    Loop
    .Rows = .Rows + 1
    .Row = 1
    Set RsQry = Nothing
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And SeqMst.HCode=HedMst.HCode And HedMst.HType=0 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    Set RsQry = Nothing
    RsQry.Open "Select HCode,IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) As Amt From TmpTrialBal Where AcCode In (" & mAcList & _
    ") And HSide='D' And LCode<>107 Group By HCode Order By HCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        i = .FindRow(RsQry.Fields("HCode"), 1, 6)
        If i > 0 Then .TextMatrix(i, 2) = RsQry.Fields("Amt")
        RsQry.MoveNext
    Loop
    Set RsQry = Nothing
    RsQry.Open "Select HCode,IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) AS Amt From TmpTrialBal Where AcCode In (" & mAcList & _
    ") And HSide='C' And LCode<>108 Group By HCode Order By HCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        i = .FindRow(RsQry.Fields("HCode"), 1, 7)
        If i > 0 Then .TextMatrix(i, 5) = RsQry.Fields("Amt")
        RsQry.MoveNext
    Loop
    Dim mTotal As Double
    mTotal = Round(SetProfitAll(mAcList), 2)
    Set RsQry = Nothing
    .Row = .Rows - 1
    If mTotal > 0 Then
        .TextMatrix(.Row, 0) = "Surplus carried over to Balance Sheet"
       .TextMatrix(.Row, 2) = Format(CStr(mTotal), "0.00")
    Else
        .TextMatrix(.Row, 3) = "Deficit carried over to Balance Sheet"
        .TextMatrix(.Row, 5) = Format(CStr(Abs(mTotal)), "0.00")
    End If
    .Refresh
    Set RsQry = Nothing
    RsQry.Open "Select * From NtDtl Where AcCode=" & mAcCode & " And RType=1", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        If RsQry.Fields("RSide") = "D" Then
            i = .FindRow(RsQry.Fields("HCode"), 1, 6)
            If i > 0 Then .TextMatrix(i, 1) = RsQry.Fields("Note")
        Else
            i = .FindRow(RsQry.Fields("HCode"), 1, 7)
            If i > 0 Then .TextMatrix(i, 4) = RsQry.Fields("Note")
        End If
        RsQry.MoveNext
    Loop
End With
SetFinalTot
End Sub
Private Sub NotePrintIE()
Dim i As Double
Dim mAcCodeAll As String
Dim mAcName As String
Dim mGroup As Double
Dim mGroupNm As String
Dim mTotal As Double
Dim RsDataO As New ADODB.Recordset
Dim RsDataD As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim m As Double
DbDataDB.BeginTrans
mAcCodeAll = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName='Trust (Local Fund) Account'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName='Foreign Contribution Account'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcType=5 And AcName Not In ('Trust (Local Fund) Account','Foreign Contribution Account')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcType<>5 And AcName Not In ('" & LsvClient.TextMatrix(LsvClient.Row, 0) & _
"','Trust (Local Fund) Account','Foreign Contribution Account')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
With VsfHelp
    i = 1
    Do While i < .Rows - 1
        Do While Len(mAcList) > 0
            If InStr(1, mAcList, ",") - 1 > 0 Then
                mAcCode = Val(Mid(mAcList, 1, InStr(1, mAcList, ",") - 1))
                mAcList = Mid(mAcList, InStr(1, mAcList, ",") + 1)
            Else
                mAcCode = Val(mAcList)
                mAcList = ""
            End If
            If RsClient.BOF = False Then
                RsClient.MoveFirst
                RsClient.Find "AcCode=" & mAcCode
            End If
            If RsClient.EOF = False Then
                If RsClient.Fields("AcName") = LsvClient.TextMatrix(LsvClient.Row, 0) Then
                    mAcName = Space(3) + LsvClient.TextMatrix(LsvClient.Row, 0) + IIf(Len(RsClient.Fields("City")) <> 0, (", " + RsClient.Fields("City")), "")
                ElseIf RsClient.Fields("AcName") = "Trust Account" Then
                    mAcName = Space(2) + RsClient.Fields("AcName")
                ElseIf RsClient.Fields("AcName") = "Foreign Contribution Account" Then
                    mAcName = Space(1) + RsClient.Fields("AcName")
                Else
                    mAcName = RsClient.Fields("AcName") + IIf(Len(RsClient.Fields("City")) <> 0, (", " + RsClient.Fields("City")), "")
                End If
            End If
            i = 1
            Do While i < .Rows - 1
                If Val(.TextMatrix(i, 1)) <> 0 Then
                    If Mid(.TextMatrix(i, 0), 1, 1) = "(" Then
                        m = i
                        Do While m > 0
                            If Mid(.TextMatrix(m, 0), 1, 1) <> "(" Then Exit Do Else m = m - 1
                        Loop
                        mGroupNm = .TextMatrix(m, 0)
                    Else
                        mGroupNm = .TextMatrix(i, 0)
                    End If
                    mGroup = Val(.TextMatrix(i, 6))
                    Set RsDataO = Nothing
                    RsDataO.Open "Select EName,IIF(Side=HSide,Amt,Amt*-1) As Amount From TmpCtDtl Where AcCode=" & mAcCode & " And HCode=" & mGroup & " And LCode Not In (107,108) Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    Do While RsDataO.EOF = False
                        If mGroupNm = .TextMatrix(i, 0) Then
                            DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & _
                            "','" & IIf(RsDataO.Fields("Amount") >= 0, Space(1) + RsDataO.Fields("EName"), "Less: " + RsDataO.Fields("EName")) & "'," & RsDataO.Fields("Amount") & ",0,'" & mAcName & "')"
                        Else
                            DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & mGroupNm & " - " & Mid(.TextMatrix(i, 0), 5) & _
                            "','" & IIf(RsDataO.Fields("Amount") >= 0, Space(1) + RsDataO.Fields("EName"), "Less: " + RsDataO.Fields("EName")) & "'," & RsDataO.Fields("Amount") & ",0,'" & mAcName & "')"
                        End If
                        RsDataO.MoveNext
                    Loop
                    i = i + 1
                    If i = .Rows - 1 Then Exit Do
                Else
                    i = i + 1
                    If i = .Rows - 1 Then Exit Do
                End If
            Loop
        Loop
    Loop
    i = 1
    mAcList = mAcCodeAll
    Do While i < .Rows - 1
        Do While Len(mAcList) > 0
            If InStr(1, mAcList, ",") - 1 > 0 Then
                mAcCode = Val(Mid(mAcList, 1, InStr(1, mAcList, ",") - 1))
                mAcList = Mid(mAcList, InStr(1, mAcList, ",") + 1)
            Else
                mAcCode = Val(mAcList)
                mAcList = ""
            End If
            If RsClient.BOF = False Then
                RsClient.MoveFirst
                RsClient.Find "AcCode=" & mAcCode
            End If
            If RsClient.EOF = False Then
                If RsClient.Fields("AcName") = LsvClient.TextMatrix(LsvClient.Row, 0) Then
                    mAcName = Space(3) + LsvClient.TextMatrix(LsvClient.Row, 0) + IIf(Len(RsClient.Fields("City")) <> 0, (", " + RsClient.Fields("City")), "")
                ElseIf RsClient.Fields("AcName") = "Trust Account" Then
                    mAcName = Space(2) + RsClient.Fields("AcName")
                ElseIf RsClient.Fields("AcName") = "Foreign Contribution Account" Then
                    mAcName = Space(1) + RsClient.Fields("AcName")
                Else
                    mAcName = RsClient.Fields("AcName") + IIf(Len(RsClient.Fields("City")) <> 0, (", " + RsClient.Fields("City")), "")
                End If
            End If
            i = 1
            Do While i < .Rows - 1
                If Val(.TextMatrix(i, 4)) <> 0 Then
                    If Mid(.TextMatrix(i, 3), 1, 1) = "(" Then
                        m = i
                        Do While m > 0
                            If Mid(.TextMatrix(m, 3), 1, 1) <> "(" Then Exit Do Else m = m - 1
                        Loop
                        mGroupNm = .TextMatrix(m, 3)
                    Else
                        mGroupNm = .TextMatrix(i, 3)
                    End If
                    mGroup = Val(.TextMatrix(i, 7))
                    Set RsDataO = Nothing
                    RsDataO.Open "Select EName,IIF(Side=HSide,Amt,Amt*-1) As Amount From TmpCtDtl Where AcCode=" & mAcCode & " And HCode=" & mGroup & " And LCode Not In (107,108) Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    Do While RsDataO.EOF = False
                        If mGroupNm = .TextMatrix(i, 3) Then
                            DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                            "','" & IIf(RsDataO.Fields("Amount") >= 0, Space(1) + RsDataO.Fields("EName"), "Less: " + RsDataO.Fields("EName")) & "'," & RsDataO.Fields("Amount") & ",0,'" & mAcName & "')"
                        Else
                            DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & mGroupNm & " - " & Mid(.TextMatrix(i, 3), 5) & _
                            "','" & IIf(RsDataO.Fields("Amount") >= 0, Space(1) + RsDataO.Fields("EName"), "Less: " + RsDataO.Fields("EName")) & "'," & RsDataO.Fields("Amount") & ",0,'" & mAcName & "')"
                        End If
                        RsDataO.MoveNext
                    Loop
                    i = i + 1
                    If i = .Rows - 1 Then Exit Do
                Else
                    i = i + 1
                    If i = .Rows - 1 Then Exit Do
                End If
            Loop
        Loop
    Loop
End With
DbDataDB.CommitTrans
mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
SetParent
End Sub
Private Sub Schedule9C()
Dim i As Double
Dim RsClient As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
With VsfHelp
    .Cols = 4
    .Rows = 1
    .TextMatrix(0, 0) = "SR."
    .ColWidth(0) = 400
    .TextMatrix(0, 1) = "SCHEDULE 9C ITEM"
    .ColWidth(1) = 7500
    .TextMatrix(0, 2) = "SUB TOTAL"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "FINAL TOTAL"
    .ColWidth(3) = 1500
    .ColFormat(3) = "0.00"
    .Refresh
    .Rows = 2
    .Row = 1
    .TextMatrix(.Row, 0) = "1"
    .TextMatrix(.Row, 1) = "Gross Annual Income"
    RowInc
    .TextMatrix(.Row, 1) = "Details of income not chargeable to contribution under Section 58 and Rule 32:"
    RowInc
    .TextMatrix(.Row, 1) = "(i) Donations received during the year from any source"
    RowInc
    .TextMatrix(.Row, 1) = "     (a) Corpus"
    .TextMatrix(.Row, 0) = "2"
    RowInc
    .TextMatrix(.Row, 1) = "          (1) From Country"
    RowInc
    .TextMatrix(.Row, 0) = "3"
    .TextMatrix(.Row, 1) = "          (2) From Foreign Country (FC)"
    RowInc
    .TextMatrix(.Row, 1) = "     (b) General"
    RowInc
    .TextMatrix(.Row, 0) = "4"
    .TextMatrix(.Row, 1) = "          (1) From Country"
    RowInc
    .TextMatrix(.Row, 0) = "5"
    .TextMatrix(.Row, 1) = "          (2) From Foreign Country (FC)"
    RowInc
    .TextMatrix(.Row, 1) = "(ii) Grants by Government and Local Authorities"
    RowInc
    .TextMatrix(.Row, 0) = "6"
    .TextMatrix(.Row, 1) = "     (a) Government and Local Authorities"
    RowInc
    .TextMatrix(.Row, 0) = "7"
    .TextMatrix(.Row, 1) = "     (b) From Foreign Country"
    RowInc
    .TextMatrix(.Row, 1) = "     (c) By Funding Agencies"
    RowInc
    .TextMatrix(.Row, 0) = "8"
    .TextMatrix(.Row, 1) = "          (1) From Country"
    RowInc
    .TextMatrix(.Row, 0) = "9"
    .TextMatrix(.Row, 1) = "          (2) From Foreign Country (FC)"
    RowInc
    .TextMatrix(.Row, 0) = "10"
    .TextMatrix(.Row, 1) = "(iii) Amount spent for the purpose of Education"
    RowInc
    .TextMatrix(.Row, 0) = "11"
    .TextMatrix(.Row, 1) = "(iv) Amount spent for the purpose of Medical Relief"
    RowInc
    .TextMatrix(.Row, 1) = "(v) (A) Deductions out of incomes from lands used for Agricultural Purpose"
    RowInc
    .TextMatrix(.Row, 0) = "12"
    .TextMatrix(.Row, 1) = "          (a) Land Revenue and Local Fund Cess"
    RowInc
    .TextMatrix(.Row, 0) = "13"
    .TextMatrix(.Row, 1) = "          (b) Rent payable to Superior Landlord"
    RowInc
    .TextMatrix(.Row, 0) = "14"
    .TextMatrix(.Row, 1) = "          (c) Cost of Production, if lands are cultivated by the Trust"
    RowInc
    .TextMatrix(.Row, 0) = "15"
    .TextMatrix(.Row, 1) = "     (B) Income from lands used for Agricultural Purpose"
    RowInc
    .TextMatrix(.Row, 1) = "(vi) (A) Deductions out of income from lands used for Non-agricultural Purpose"
    RowInc
    .TextMatrix(.Row, 0) = "16"
    .TextMatrix(.Row, 1) = "          (a) Assessment, Cesses and other Municipal Taxes"
    RowInc
    .TextMatrix(.Row, 0) = "17"
    .TextMatrix(.Row, 1) = "          (b) Ground Rent payable to Superior Landlord"
    RowInc
    .TextMatrix(.Row, 0) = "18"
    .TextMatrix(.Row, 1) = "          (c) Insurance Premium"
    RowInc
    .TextMatrix(.Row, 0) = "19"
    .TextMatrix(.Row, 1) = "          (d) Repairs at 8.33% of gross rent of building"
    RowInc
    .TextMatrix(.Row, 0) = "20"
    .TextMatrix(.Row, 1) = "          (e) Cost of collection at 4% of gross rent of buildings let-out"
    RowInc
    .TextMatrix(.Row, 0) = "21"
    .TextMatrix(.Row, 1) = "     (B) Income from lands used for Non-agricultural Purpose"
    RowInc
    .TextMatrix(.Row, 0) = "22"
    .TextMatrix(.Row, 1) = "(vii) Cost of collection of income or receipts from securities, stocks, etc. at 1% of such income"
    RowInc
    .TextMatrix(.Row, 0) = "23"
    .TextMatrix(.Row, 1) = "(viii) Deductions on account of repairs in respect of buildings not rented and not yielding income at 8.33% of the estimated gross annual rent"
    RowInc
    .TextMatrix(.Row, 0) = "24"
    .TextMatrix(.Row, 1) = "Income Liable to Contribution"
    .Editable = flexEDNone
    .AutoSearch = flexSearchFromCursor
    RsQ.Open "Select * From 9CDtl Where AcCode=" & mAcCode & " Order By SrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.EOF = False Then
        .Row = 1
        Do While RsQ.EOF = False
            .TextMatrix(.Row, 2) = RsQ.Fields("Amt1")
            .TextMatrix(.Row, 3) = RsQ.Fields("Amt2")
            RsQ.MoveNext
            If RsQ.EOF = False Then
                If .Row + 1 < .Rows Then .Row = .Row + 1
            End If
        Loop
        TxtCTotal.Text = "0.00"
        TxtCTotal.Text = Val(.TextMatrix(1, 3))
        For i = 2 To .Rows - 1
            TxtCTotal.Text = Val(TxtCTotal.Text) - Val(.TextMatrix(i, 3))
            If Val(.TextMatrix(i, 2)) = 0 Then .TextMatrix(i, 2) = ""
            If Val(.TextMatrix(i, 3)) = 0 Then .TextMatrix(i, 3) = ""
        Next
        If Val(TxtCTotal.Text) < 0 Then TxtCTotal.Text = "0.00"
    Else
        MsgBox "Schedule 9C not generated. Please generate data and print again.", vbCritical, "Alert"
        Exit Sub
    End If
End With
RsClient.Open "Select * From AcMst Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsClient.EOF = False Then
    With RepPrint
        .Connect = MSCONNECT
        .ReportFileName = App.Path + "\Report\A9CRpt.Rpt"
        .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
        .Formulas(1) = "mNAInc='" & Format(VsfHelp.TextMatrix(29, 2), "0.00") & "'"
        .Formulas(2) = "mSecColl='" & Format(VsfHelp.TextMatrix(30, 3), "0.00") & "'"
        .Formulas(3) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
        .Formulas(4) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
        .Formulas(5) = "mNAColl='" & Format(VsfHelp.TextMatrix(28, 2), "0.00") & "'"
        .Formulas(6) = "mNATotal='" & Format(VsfHelp.TextMatrix(28, 3), "0.00") & "'"
        .Formulas(7) = "mNARep='" & Format(VsfHelp.TextMatrix(27, 2), "0.00") & "'"
        .Formulas(8) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
        .Formulas(9) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
        .Formulas(10) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
        .Formulas(11) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
        .Formulas(12) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
        .Formulas(13) = "mNAIns='" & Format(VsfHelp.TextMatrix(26, 3), "0.00") & "'"
        .Formulas(14) = "mRepNR='" & Format(VsfHelp.TextMatrix(31, 3), "0.00") & "'"
        .Formulas(15) = "mTotal='" & Format(TxtCTotal.Text, "0.00") & "'"
        .Formulas(16) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
        .Formulas(17) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
        .Formulas(18) = "mGrossInc='" & Format(VsfHelp.TextMatrix(1, 3), "0.00") & "'"
        .Formulas(19) = "mDCLocal='" & Format(VsfHelp.TextMatrix(5, 2), "0.00") & "'"
        .Formulas(20) = "mDCTotal='" & Format(VsfHelp.TextMatrix(6, 3), "0.00") & "'"
        .Formulas(21) = "mGCLocal='" & Format(VsfHelp.TextMatrix(8, 2), "0.00") & "'"
        .Formulas(22) = "mGCFC='" & Format(VsfHelp.TextMatrix(9, 2), "0.00") & "'"
        .Formulas(23) = "mGCTotal='" & Format(VsfHelp.TextMatrix(9, 3), "0.00") & "'"
        .Formulas(24) = "mGGovt='" & Format(VsfHelp.TextMatrix(11, 3), "0.00") & "'"
        .Formulas(25) = "mGFC='" & Format(VsfHelp.TextMatrix(12, 3), "0.00") & "'"
        .Formulas(26) = "mGFLocal='" & Format(VsfHelp.TextMatrix(14, 2), "0.00") & "'"
        .Formulas(27) = "mGFFC='" & Format(VsfHelp.TextMatrix(15, 2), "0.00") & "'"
        .Formulas(28) = "mGFTotal='" & Format(VsfHelp.TextMatrix(15, 3), "0.00") & "'"
        .Formulas(29) = "mExEdu='" & Format(VsfHelp.TextMatrix(16, 3), "0.00") & "'"
        .Formulas(30) = "mExMed='" & Format(VsfHelp.TextMatrix(17, 3), "0.00") & "'"
        .Formulas(31) = "mAgLRev='" & Format(VsfHelp.TextMatrix(19, 2), "0.00") & "'"
        .Formulas(32) = "mAgRent='" & Format(VsfHelp.TextMatrix(20, 2), "0.00") & "'"
        .Formulas(33) = "mAgExp='" & Format(VsfHelp.TextMatrix(21, 2), "0.00") & "'"
        .Formulas(34) = "mAgTotal='" & Format(VsfHelp.TextMatrix(21, 3), "0.00") & "'"
        .Formulas(35) = "mAgInc='" & Format(VsfHelp.TextMatrix(22, 2), "0.00") & "'"
        .Formulas(36) = "mNATax='" & Format(VsfHelp.TextMatrix(24, 2), "0.00") & "'"
        .Formulas(37) = "mNARent='" & Format(VsfHelp.TextMatrix(25, 2), "0.00") & "'"
        .Formulas(38) = "mTName='" & LsvClient.TextMatrix(LsvClient.Row, 0) & ", " & LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
        .Formulas(39) = "mTAddress='" & RsClient.Fields("Address") & IIf(Len(RsClient.Fields("Address")) > 0, ", ", "") & RsClient.Fields("City") & IIf(RsClient.Fields("Taluka") <> RsClient.Fields("City"), ", " + RsClient.Fields("Taluka"), "") & IIf((RsClient.Fields("District") <> RsClient.Fields("Taluka") And RsClient.Fields("District") <> RsClient.Fields("City")), ", " + RsClient.Fields("District"), "") & ", " & RsClient.Fields("State") & ", " & RsClient.Fields("PinCode") & "'"
        .Formulas(40) = "mBBank='" & RsClient.Fields("BBank") & "'"
        .Formulas(41) = "mBBranch='" & RsClient.Fields("BBranch") & "'"
        .Formulas(42) = "mBAddress='" & RsClient.Fields("BAddress") & "'"
        .Formulas(43) = "mTrustee='" & RsRep.Fields("RTrustee") & "'"
        .Formulas(44) = "mTRegNo='" & RsClient.Fields("TRegNo") & "'"
        .Formulas(45) = "mTRegDate='" & RsClient.Fields("TRegDt") & "'"
        .Formulas(46) = "mDcFc='" & Format(VsfHelp.TextMatrix(7, 2), "0.00") & "'"
        .Formulas(47) = "mTitle9c='Statement of Income liable to Contribution for the year ended on " & RsRep.Fields("RTDt") & "'"
         If Len(Trim(RsClient.Fields("FCRegNo"))) <> 0 Then
            .Formulas(48) = "mTFCBank='" & RsClient.Fields("FCBank") & ", " & RsClient.Fields("FCAcType") & " A/c No. " & RsClient.Fields("FCAcNo") & "'"
            .Formulas(49) = "mTFCNo='" & RsClient.Fields("FCRegNo") & "'"
            .Formulas(50) = "mTFCDate='" & RsClient.Fields("FCRegDt") & "'"
        Else
            .Formulas(48) = "mTFCBank='N/A'"
            .Formulas(49) = "mTFCNo='N/A'"
            .Formulas(50) = "mTFCDate='N/A'"
        End If
'        If RsClient.Fields("ObjEdu") = -1 And RsClient.Fields("ObjMed") = -1 Then
'            .Formulas(51) = "m9CNote='Education and Medical Relief'"
'        ElseIf RsClient.Fields("ObjEdu") = -1 Then
'            .Formulas(51) = "m9CNote='Education'"
'        ElseIf RsClient.Fields("ObjMed") = -1 Then
'            .Formulas(51) = "m9CNote='Medical Relief'"
'        Else
            .Formulas(51) = "m9CNote=''"
'        End If
        .Formulas(52) = "mDCFC='" & Format(VsfHelp.TextMatrix(6, 2), "0.00") & "'"
        .Formulas(53) = "mStateAct='" & RsState.Fields("RPAct") & "'"
        .Formulas(54) = "mActTitle='" & IIf(Len(RsState.Fields("IXCTitle")) > 0, RsState.Fields("IXCTitle"), "") & "'"
        .Formulas(55) = "mActSub='" & IIf(Len(RsState.Fields("IXCTitle")) > 0, RsState.Fields("IXCSub"), "") & "'"
        .Action = 1
        For i = 0 To 55
            .Formulas(i) = ""
        Next
    End With
End If
End Sub
Private Sub RowInc()
With VsfHelp
    .Rows = .Rows + 1
    .Row = .Rows - 1
End With
End Sub

Private Sub PrintData()
Dim i As Double
Dim mAcCodeAll As String
Dim RsQ As New ADODB.Recordset
mAcCodeAll = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName='Trust (Local Fund) Account'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName='Foreign Contribution Account'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName Not In ('" & LsvClient.TextMatrix(LsvClient.Row, 0) & _
"','Trust (Local Fund) Account','Foreign Contribution Account')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
'Audit Report
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\AdtRpt.Rpt"
    .SelectionFormula = "{TranBr.RepName}='VPDUW\'"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(2) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(3) = "mTitle1='Independent Auditor`s Report'"
    .Formulas(4) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(5) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
    .Formulas(6) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(7) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(8) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(9) = "mTRegNo='" & RsClient.Fields("TRegNo") & "'"
    .Formulas(10) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(11) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(12) = "mTName='" & LsvClient.TextMatrix(LsvClient.Row, 0) & ", " & LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Formulas(13) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(14) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(15) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(16) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(17) = "mCYear='" & CStr(Year(RsRep.Fields("RtDt"))) & "'"
    If ComQualify.Text = "Yes" Then .Formulas(18) = "mQualif='Qualified Opinion'" Else .Formulas(18) = "mQualif='Opinion'"
    .Formulas(19) = "mQulRpt='" & TxtQualify.Text & "'"
    .Formulas(20) = "mEOMRpt='" & TxtEmphasis.Text & "'"
    .Formulas(21) = "mAOMRpt='" & TxtOMetter.Text & "'"
    If ComAcMethod.Text = "Cash Basis" Then .Formulas(22) = "mAcBase='cash basis'" Else .Formulas(22) = "mAcBase='accrual basis'"
    If ComFAsset.Text = "Prepared" Then .Formulas(23) = "mFAInvt='An'" Else .Formulas(23) = "mFAInvt='No'"
    .Formulas(24) = "mCType='T'"
    .Formulas(25) = "mStateAct='" & RsState.Fields("RPAct") & "'"
    .Formulas(26) = "mStateActNo='" & IIf(Len(RsState.Fields("RPActNo")) > 0, RsState.Fields("RPActNo"), "") & "'"
    .Formulas(27) = "mStateCd='" & RsState.Fields("StateCd") & "'"
    .Action = 1
    For i = 0 To 27
        .Formulas(i) = ""
    Next
    .SelectionFormula = ""
End With
'Balance Sheet
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\ABSRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(2) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(3) = "mTitle1='Balance Sheet as on " & RsRep.Fields("RtDt") & "'"
    .Formulas(4) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(5) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
    .Formulas(6) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(7) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
'    .Formulas(8) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(9) = "mTRegNo='" & RsClient.Fields("TRegNo") & "'"
    .Formulas(10) = "mTRegDate='" & RsClient.Fields("TRegDt") & "'"
    .Formulas(11) = "mTAddress='" & RsClient.Fields("Address") & IIf(Len(RsClient.Fields("Address")) > 0, ", ", "") & RsClient.Fields("City") & IIf(RsClient.Fields("Taluka") <> RsClient.Fields("City"), ", " + RsClient.Fields("Taluka"), "") & IIf((RsClient.Fields("District") <> RsClient.Fields("Taluka") And RsClient.Fields("District") <> RsClient.Fields("City")), ", " + RsClient.Fields("District"), "") & ", " & RsClient.Fields("State") & ", " & RsClient.Fields("PinCode") & "'"
    If Len(Trim(RsClient.Fields("FCRegNo"))) <> 0 Then
        .Formulas(12) = "mTFCBank='" & RsClient.Fields("FCBank") & ", " & RsClient.Fields("FCAcType") & " A/c No.: " & RsClient.Fields("FCAcNo") & "'"
        .Formulas(13) = "mTFCNo='" & RsClient.Fields("FCRegNo") & "'"
        .Formulas(14) = "mTFCDate='" & RsClient.Fields("FCRegDt") & "'"
    Else
        .Formulas(12) = "mTFCBank='N/A'"
        .Formulas(13) = "mTFCNo='N/A'"
        .Formulas(14) = "mTFCDate='N/A'"
    End If
    .Formulas(15) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(16) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(17) = "mTName='" & LsvClient.TextMatrix(LsvClient.Row, 0) & ", " & LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Formulas(18) = "mHOff='" & RsComp.Fields("Add1") & ", " & RsComp.Fields("Add2") & ", " & RsComp.Fields("City") & "'"
    .Formulas(19) = "mStateAct='" & RsState.Fields("RPAct") & "'"
    .Formulas(20) = "mActTitle='" & IIf(Len(RsState.Fields("BSTitle")) > 0, RsState.Fields("BSTitle"), "") & "'"
    .Formulas(21) = "mActSub='" & IIf(Len(RsState.Fields("BSSub")) > 0, RsState.Fields("BSSub"), "") & "'"
    .Action = 1
    For i = 0 To 21
        .Formulas(i) = ""
    Next
End With
'P & L
If RsClient.BOF = False Then
    RsClient.MoveFirst
    RsClient.Find "AcCode=" & Val(LsvClient.TextMatrix(LsvClient.Row, 5))
End If
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\AIERpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(2) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(3) = "mTitle1='Income and Expenditure Account for the year ended on " & RsRep.Fields("RtDt") & "'"
    .Formulas(4) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(5) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
    .Formulas(6) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(7) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
'    .Formulas(8) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(9) = "mTRegNo='" & RsClient.Fields("TRegNo") & "'"
    .Formulas(10) = "mTRegDate='" & RsClient.Fields("TRegDt") & "'"
    .Formulas(11) = "mTAddress='" & RsClient.Fields("Address") & IIf(Len(RsClient.Fields("Address")) > 0, ", ", "") & RsClient.Fields("City") & IIf(RsClient.Fields("Taluka") <> RsClient.Fields("City"), ", " + RsClient.Fields("Taluka"), "") & IIf((RsClient.Fields("District") <> RsClient.Fields("Taluka") And RsClient.Fields("District") <> RsClient.Fields("City")), ", " + RsClient.Fields("District"), "") & ", " & RsClient.Fields("State") & ", " & RsClient.Fields("PinCode") & "'"
    If Len(Trim(RsClient.Fields("FCRegNo"))) <> 0 Then
        .Formulas(12) = "mTFCBank='" & RsClient.Fields("FCBank") & ", " & RsClient.Fields("FCAcType") & " A/c No.: " & RsClient.Fields("FCAcNo") & "'"
        .Formulas(13) = "mTFCNo='" & RsClient.Fields("FCRegNo") & "'"
        .Formulas(14) = "mTFCDate='" & RsClient.Fields("FCRegDt") & "'"
    Else
        .Formulas(12) = "mTFCBank='N/A'"
        .Formulas(13) = "mTFCNo='N/A'"
        .Formulas(14) = "mTFCDate='N/A'"
    End If
    .Formulas(15) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(16) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(17) = "mTName='" & LsvClient.TextMatrix(LsvClient.Row, 0) & ", " & LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Formulas(18) = "mHOff='" & RsComp.Fields("Add1") & ", " & RsComp.Fields("Add2") & ", " & RsComp.Fields("City") & "'"
    .Formulas(19) = "mStateAct='" & RsState.Fields("RPAct") & "'"
    .Formulas(20) = "mActTitle='" & IIf(Len(RsState.Fields("IETitle")) > 0, RsState.Fields("IETitle"), "") & "'"
    .Formulas(21) = "mActSub='" & IIf(Len(RsState.Fields("IESub")) > 0, RsState.Fields("IESub"), "") & "'"
    .Action = 1
    For i = 0 To 21
        .Formulas(i) = ""
    Next
End With
'Note
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\NteRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(9) = "mTitle1='Notes to Financial Statements'"
    .Formulas(10) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(11) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(12) = "mTrack='" & CStr(InStr(1, mAcCodeAll, ",")) & "'"
    .Formulas(13) = "mType=''"
    .Formulas(14) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Action = 1
    For i = 0 To 14
        .Formulas(i) = ""
    Next
End With
mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
SetParent
End Sub
Private Function SaveData()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From ArpDtl Where AcCode=" & mAcCode
DbDataDB.Execute "Insert InTo ArpDtl (AcCode,AcBase,FAInvt,Qualif,QulRpt,EOMRpt,AOMRpt) Values (" & mAcCode & "," & ComAcMethod.ItemData(ComAcMethod.ListIndex) & "," & _
IIf(ComFAsset.ItemData(ComFAsset.ListIndex) = 1, 0, -1) & "," & IIf(ComQualify.ItemData(ComQualify.ListIndex) = 1, 0, -1) & ",'" & TxtQualify.Text & "','" & _
TxtEmphasis.Text & "','" & TxtOMetter.Text & "')"
'DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'AUDIT_RPT','SAVE_DATA','" & Date & "','" & Time & "')"
DbDataDB.CommitTrans
MsgBox "Record Saved Successfully.", vbInformation, "Alert"
Exit Function
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Function
Private Sub TxtQualify_GotFocus()
    If ComQualify.Text = "Yes" Then TxtQualify.Locked = False Else TxtQualify.Locked = True
End Sub

Private Sub Display()
Dim RsQ As New ADODB.Recordset
RsQ.Open "Select * From ArpDtl Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With RsQ
    If .EOF = True Then Exit Sub
    ComAcMethod.ListIndex = IIf(.Fields("AcBase") = 2, 1, 0)
    ComFAsset.ListIndex = IIf(.Fields("FAInvt") = -1, 1, 0)
    ComQualify.ListIndex = IIf(.Fields("Qualif") = -1, 1, 0)
    TxtQualify.Text = IIf(IsNull(.Fields("QulRpt")) = False, .Fields("QulRpt"), "")
    TxtEmphasis.Text = IIf(IsNull(.Fields("EOMRpt")) = False, .Fields("EOMRpt"), "")
    TxtOMetter.Text = IIf(IsNull(.Fields("AOMRpt")) = False, .Fields("AOMRpt"), "")
End With
End Sub

Private Sub ClearAll()
Dim ObjText As Object
For Each ObjText In Me
    If TypeOf ObjText Is TextBox Then ObjText.Text = ""
Next
mAcCode = 0
End Sub
