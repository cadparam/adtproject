VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmBS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sheet"
   ClientHeight    =   9570
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   15765
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FBSh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   15765
   Begin VB.Frame FraNote 
      Height          =   6612
      Left            =   18000
      TabIndex        =   11
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         Picture         =   "FBSh.frx":000C
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
         FormatString    =   $"FBSh.frx":044E
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
      TabIndex        =   9
      Top             =   1920
      Width           =   14750
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   3852
         Left            =   120
         TabIndex        =   10
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
         FormatString    =   $"FBSh.frx":0497
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
      TabIndex        =   7
      Top             =   0
      Width           =   15615
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
         TabIndex        =   26
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   855
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
         Height          =   372
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   1584
      End
      Begin VB.Frame FraMainExp 
         BackColor       =   &H00F7E0FE&
         Height          =   5412
         Left            =   1.80000e5
         TabIndex        =   21
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
            TabIndex        =   22
            ToolTipText     =   "Cancel"
            Top             =   4920
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfMainExport 
            Height          =   5052
            Left            =   0
            TabIndex        =   23
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
            FormatString    =   $"FBSh.frx":04E0
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
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   9000
         Width           =   1776
      End
      Begin VB.Frame FraDetail 
         Height          =   6132
         Left            =   18000
         TabIndex        =   15
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
            TabIndex        =   18
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
            Picture         =   "FBSh.frx":0529
            TabIndex        =   17
            ToolTipText     =   "Cancel"
            Top             =   5640
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfDetail 
            Height          =   5412
            Left            =   120
            TabIndex        =   16
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
            FormatString    =   $"FBSh.frx":096B
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
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   855
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
         Picture         =   "FBSh.frx":09B4
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
         TabIndex        =   6
         Top             =   240
         Width           =   4896
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   8200
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   15255
         _cx             =   26908
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
         FormatString    =   $"FBSh.frx":0FEA
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
         Picture         =   "FBSh.frx":1033
         TabIndex        =   4
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   800
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
         TabIndex        =   24
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
         TabIndex        =   8
         Top             =   240
         Width           =   1236
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   9375
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   15615
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
Attribute VB_Name = "FrmBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim mAcCode As Double
Dim mProfit As Double
Dim RsLedger As New ADODB.Recordset
Private Sub CmdSearch_Click()
    FraClientHelp.Left = 180
    LsvClient.SetFocus
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
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode,'' As RName,AcMst.PACode From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 And " & _
"AcMst.AcType=GrpMst.GCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.City,AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,''),GrpMst.GName,AcMst.AcCode," & _
"'',AcMst.PACode From AcMst,GrpMst,AcMst As AcMst1 Where AcMst.Active=-1 And AcMst.PaCode=AcMst1.AcCode And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    .TextMatrix(0, 6) = "MAIN PARENT"
    .ColWidth(6) = 4000
    .Col = 1
    .ExtendLastCol = True
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
    Set RsQry = Nothing
    RsQry.Open "Select AcMst.AcCode,AcMst2.AcName+', '+AcMst2.City As RName From AcMst,AcMst As AcMst1,AcMst As AcMst2 Where AcMst.PaCode=AcMst1.AcCode And " & _
    "AcMst1.PaCode=AcMst2.AcCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .Row = 1
        .Row = .FindRow(RsQry.Fields("AcCode"), , 5)
        If .Row > 1 Then .TextMatrix(.Row, 6) = RsQry.Fields("RName")
        RsQry.MoveNext
    Loop
End With
End Function
Private Sub ClearText()
    TxtName.Text = ""
    mAcCode = 0
    FraClientHelp.Left = 18000
    FraDetail.Left = 18000
    FraNote.Left = 18000
End Sub
Private Function SaveNote()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete NtDtl.* From NtDtl,HedMst Where NtDtl.RType=2 And NtDtl.AcCode=" & mAcCode & " And NtDtl.HCode=HedMst.HCode And HedMst.HType=1"
With VsfHelp
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 And .TextMatrix(i, 0) <> "" Then
            DbDataDB.Execute "Insert InTo NtDtl (AcCode,HCode,[Note],RSide,RType) Values (" & mAcCode & "," & Val(VsfHelp.TextMatrix(i, 6)) & "," & Val(.TextMatrix(i, 1)) & ",'C',2)"
        End If
    Next
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 4)) <> 0 And .TextMatrix(i, 3) <> "" Then
            DbDataDB.Execute "Insert InTo NtDtl (AcCode,HCode,[Note],RSide,RType) Values (" & mAcCode & "," & Val(VsfHelp.TextMatrix(i, 7)) & "," & Val(.TextMatrix(i, 4)) & ",'D',2)"
        End If
    Next
End With
DbDataDB.CommitTrans
MsgBox "Record saved.", vbInformation, "Alert"
Exit Function
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDIWork.PctMdi.Visible = True
    Set DBWorkTmp = Nothing
    Unload Me
End Sub

Private Sub LsvClient_KeyPress(KeyAscii As Integer)
Dim RsT As New ADODB.Recordset
Dim mTName As String
If KeyAscii = 13 Then
    TxtName.Text = LsvClient.TextMatrix(LsvClient.Row, 0)
    mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
    FraClientHelp.Left = 18000
    mProfit = 0
    Dim RsQ As New ADODB.Recordset
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'BSHEET_IND','VIEW','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    DBWorkTmp.BeginTrans
    DBWorkTmp.Execute "Delete From TmpSMPrn"
    DBWorkTmp.Execute "Delete From TmpTrialBal"
    DBWorkTmp.Execute "Delete From TmpCtDtl"
    DBWorkTmp.Execute "Delete From TmpRpDtl"
    DBWorkTmp.CommitTrans
    DBWorkTmp.BeginTrans
    Set RsQ = Nothing
    RsQ.Open "Select * From QTrialBal Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpTrialBal (HType,HSide,AcCode,HCode,LCode,OpDr,OpCr,ADr,ACr,DBal,CBal) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & _
        "'," & RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("OpDr") & "," & RsQ.Fields("OpCr") & "," & RsQ.Fields("ADr") & _
        "," & RsQ.Fields("ACr") & "," & RsQ.Fields("DBal") & "," & RsQ.Fields("CBal") & ")"
        RsQ.MoveNext
    Loop
    
    Set RsQ = Nothing
    RsQ.Open "Select * From QCtDtl Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        If (RsQ.Fields("ECode") = 19 Or RsQ.Fields("ECode") = 20) And RsQ.Fields("TrfCode") <> 0 Then
            Set RsT = Nothing
            RsT.Open "Select * From AcMst Where AcCode=" & RsQ.Fields("TrfCode"), DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsQ.Fields("Side") = "C" Then
                mTName = "Transfer from " & IIf(IsNull(RsT.Fields("City")) = False, RsT.Fields("AcName") & ", " & RsT.Fields("City"), RsT.Fields("AcName"))
            Else
                mTName = "Transfer to " & IIf(IsNull(RsT.Fields("City")) = False, RsT.Fields("AcName") & ", " & RsT.Fields("City"), RsT.Fields("AcName"))
            End If
            DBWorkTmp.Execute "Insert InTo TmpCtDtl (HType,HSide,AcCode,HCode,LCode,ECode,EName,Side,Amt,TrfCode) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
            RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & mTName & "','" & RsQ.Fields("Side") & _
            "'," & RsQ.Fields("Amt") & "," & RsQ.Fields("TrfCode") & ")"
        Else
            DBWorkTmp.Execute "Insert InTo TmpCtDtl (HType,HSide,AcCode,HCode,LCode,ECode,EName,Side,Amt,TrfCode) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
            RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("EName") & "','" & RsQ.Fields("Side") & _
            "'," & RsQ.Fields("Amt") & "," & RsQ.Fields("TrfCode") & ")"
        End If
    RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From QRpDtl Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpRpDtl (HType,HSide,AcCode,HCode,LCode,ECode,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
        RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("Side") & _
        "'," & RsQ.Fields("Amt") & ")"
        RsQ.MoveNext
    Loop
    DBWorkTmp.CommitTrans
    If mFinYear = "19-20" Then SetData1 Else SetData
    VsfHelp.SetFocus
End If
End Sub
Private Sub VsfHelp_DblClick()
    SetLedger VsfHelp.Col
    If VsfNote.TextMatrix(1, 0) <> "" Then
        FraNote.Left = 180
        VsfNote.SetFocus
    End If
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
                DbDataDB.BeginTrans
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'BSHEET_IND','SAVE_NOTES','" & Date & "','" & Time & "')"
                DbDataDB.CommitTrans
                VsfHelp.SetFocus
            End If
        End If
    Case "Cancel"
        FraNote.Left = 18000
        FraDetail.Left = 18000
        VsfHelp.SetFocus
    Case "Exit"
        DBWorkTmp.Close
        Unload Me
    Case "Print"
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'BSHEET_IND','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(0).Enabled = mVal
    TlbSav(1).Enabled = True
    TlbSav(2).Enabled = True
End Function
Private Sub SetData()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And SeqMst.HCode=HedMst.HCode" & _
" And HedMst.HType=1 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    RsQry.Open "Select HCode As HeadCode,Sum(CBal) As CTotRs,Sum(DBal) As DTotRs From TmpTrialBal Where HType=1 Group By HCode Order By HCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQry.EOF = False
        If RsQry.Fields("CTotRs") <> 0 Then
            i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
            If i > 0 Then
                If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 5 Then    '   Income And Expenditure
                    mProfit = SetProfit(mAcCode)
                    .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("CTotRs")) + mProfit, "0.00")
                Else
                    .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("CTotRs")), "0.00")
                End If
            End If
        Else
            i = .FindRow(RsQry.Fields("HeadCode"), 1, 7)
            If i > 0 Then
                .TextMatrix(i, 5) = Format(CStr(Val(.TextMatrix(i, 5)) + RsQry.Fields("DTotRs")), "0.00")
            End If
        End If
        RsQry.MoveNext
    Loop
    If mProfit = 0 Then mProfit = SetProfit(mAcCode)
    i = .FindRow(9, 1, 6)
    If i > 0 Then
        If Val(.TextMatrix(i, 2)) = 0 And mProfit <> 0 Then .TextMatrix(i, 2) = mProfit
        Set RsQry = Nothing
        RsQry.Open "Select Sum(IIF(JvDtl.Side='D',JvDtl.Amt*-1,JvDtl.Amt)) As RTotal From JvDtl,EntMst,LedMst Where JvDtl.AcCode=" & mAcCode & " And JvDtl.ECode=EntMst.ECode" & _
        " And EntMst.LCode=LedMst.LCode And LedMst.HCode In (58,5)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(i, 2) = Val(.TextMatrix(i, 2)) + IIf(IsNull(RsQry.Fields("RTotal")) = True, 0, RsQry.Fields("RTotal"))
        .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "0.00")
    End If
    Set RsQry = Nothing
    RsQry.Open "Select * From NtDtl Where AcCode=" & mAcCode & " And RType=2 Order By HCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
Private Sub SetLedger(ByVal mSide As Integer)
Dim RsQry As New ADODB.Recordset
Dim mRow As Double
Dim mHeadCode As Double
If mSide <= 2 Then mHeadCode = Val(VsfHelp.TextMatrix(VsfHelp.Row, 6)) Else mHeadCode = Val(VsfHelp.TextMatrix(VsfHelp.Row, 7))
Set RsLedger = Nothing
RsLedger.Open "Select LedMst.LName,LedMst.LCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & _
") And LedMst.HCode=" & mHeadCode & " And LedMst.HCode=SeqMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And SeqMst.HCode=" & _
"HedMst.HCode Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    Do While RsLedger.EOF = False
        .TextMatrix(.Row, 0) = RsLedger.Fields("LName")
        .TextMatrix(.Row, 4) = RsLedger.Fields("LCode")
        RsLedger.MoveNext
        If RsLedger.EOF = False Then
            .Rows = .Rows + 1
            .Row = .Row + 1
        End If
    Loop
    For mHeadCode = 1 To .Rows - 1
        Set RsLedger = Nothing
        RsLedger.Open "Select * From TmpTrialBal Where LCode=" & Val(.TextMatrix(mHeadCode, 4)), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsLedger.EOF = False Then
            If mSide > 2 Then
                If Mid(VsfHelp.TextMatrix(VsfHelp.Row, 3), 1, 4) = "Cash" Or Mid(VsfHelp.TextMatrix(VsfHelp.Row, 3), 1, 4) = "Bank" Then
                    .TextMatrix(mHeadCode, 3) = RsLedger.Fields("DBal")
                    .TextMatrix(mHeadCode, 1) = RsLedger.Fields("OpDr")
                Else
                    .TextMatrix(mHeadCode, 2) = RsLedger.Fields("ADR") - RsLedger.Fields("ACR")
                    .TextMatrix(mHeadCode, 1) = RsLedger.Fields("OpDr")
                    .TextMatrix(mHeadCode, 3) = RsLedger.Fields("DBal")
                End If
            Else
                If Val(.TextMatrix(mHeadCode, 4)) = 1 Or Val(.TextMatrix(mHeadCode, 4)) = 88 Then
                    .TextMatrix(mHeadCode, 1) = RsLedger.Fields("OpCr")
                    .TextMatrix(mHeadCode, 2) = SetProfit(mAcCode)
                    .TextMatrix(mHeadCode, 3) = Val(.TextMatrix(mHeadCode, 1)) + Val(.TextMatrix(mHeadCode, 2))
                Else
                    .TextMatrix(mHeadCode, 2) = RsLedger.Fields("ACR") - RsLedger.Fields("ADR")
                    .TextMatrix(mHeadCode, 1) = RsLedger.Fields("OpCr")
                    .TextMatrix(mHeadCode, 3) = RsLedger.Fields("CBal")
                End If
            End If
        Else
            If Val(.TextMatrix(mHeadCode, 4)) = 1 Then
                .TextMatrix(mHeadCode, 2) = SetProfit(mAcCode)
            ElseIf Val(.TextMatrix(mHeadCode, 4)) = 88 Then
                .TextMatrix(mHeadCode, 2) = SetProfit(mAcCode)
            End If
        End If
    Next
End With
SetTotal
End Sub

Private Sub VsfHelp_EnterCell()
With VsfHelp
    If .Col = 1 And .TextMatrix(.Row, 0) <> "" Then
        .Editable = flexEDKbd
        On Error Resume Next
        SendKeys "{F2}"
    ElseIf .Col = 4 And .TextMatrix(.Row, 3) <> "" Then
        .Editable = flexEDKbd
        On Error Resume Next
        SendKeys "{F2}"
    Else
        .Editable = flexEDNone
    End If
End With
End Sub

Private Sub VsfNote_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SetTotal
End Sub
Private Sub VsfNote_DblClick()
    If VsfNote.Col = 2 Then
        FraDetail.Left = 120
        FraNote.Left = 18000
        SetDetail
        ShowDetail
        VsfDetail.SetFocus
    End If
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
'        If Mid(VsfHelp.TextMatrix(VsfHelp.Row, 3), 1, 4) = "Cash" Or Mid(VsfHelp.TextMatrix(VsfHelp.Row, 3), 1, 4) = "Bank" Then
'
'        Else
'            .TextMatrix(i, 3) = Val(.TextMatrix(i, 1)) + Val(.TextMatrix(i, 2))
'        End If
        TxtNTotal.Text = Val(TxtNTotal.Text) + Val(.TextMatrix(i, 3))
    Next
End With
TxtOTotal.Text = Format(TxtOTotal.Text, "0.00")
TxtTotal.Text = Format(TxtTotal.Text, "0.00")
TxtNTotal.Text = Format(TxtNTotal.Text, "0.00")
End Sub
Private Sub SetDetail()
With VsfDetail
    .Rows = 1
    .Cols = 5
    .TextMatrix(0, 0) = "ACTION"
    .ColWidth(0) = 800
    .TextMatrix(0, 1) = "PARTICULAR"
    .ColWidth(1) = 4500
    .TextMatrix(0, 2) = "AMOUNT RS."
    .ColWidth(2) = 1300
    .ColAlignment(2) = flexAlignRightCenter
    .ColFormat(2) = "0.00"
    .ColWidth(3) = 100 'ECode
    .ColWidth(4) = 100 ' STATUS
    .Rows = 2
    .Row = 1
    .ColComboList(0) = "|Add|Less|"
    .Editable = flexEDKbd
    .Refresh
End With
End Sub
Private Sub VsfDetail_RowColChange()
With VsfDetail
    If .Col = 1 Then
        If .TextMatrix(.Row, 0) = "" Then
            MsgBox "Please Select Option.", vbInformation, "Alert"
            .Col = 0
            SendKeys "{F2}"
            .SetFocus
        ElseIf .TextMatrix(.Row, 0) = "Add" Then
        ElseIf .TextMatrix(.Row, 0) = "Less" Then
        Else
            MsgBox "InValid Option.", vbInformation, "Alert"
            .TextMatrix(.Row, 0) = "Add"
            .Col = 0
            .SetFocus
        End If
    End If
    If .Col <> 1 Then .Editable = flexEDKbd Else .Editable = flexEDNone
    If .Col = 3 Then
        If .Row + 1 = .Rows And (.TextMatrix(.Row, 0) <> "" And .TextMatrix(.Row, 1) <> "" And Val(.TextMatrix(.Row, 2)) <> 0) Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
        End If
    End If
    SetDTotal
End With
End Sub
Private Sub CmdCancelD_Click()
    FraDetail.Left = 18000
    FraNote.Left = 180
    VsfNote.SetFocus
End Sub
Private Sub SetDTotal()
Dim i As Double
TxtTotalD.Text = "0.00"
With VsfDetail
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) = "Add" Then TxtTotalD.Text = Val(TxtTotalD.Text) + Val(.TextMatrix(i, 2)) Else TxtTotalD.Text = Val(TxtTotalD.Text) - Val(.TextMatrix(i, 2))
    Next
End With
TxtTotalD.Text = Format(TxtTotalD.Text, "0.00")
End Sub
Private Function ShowDetail()
Dim RsQ As New ADODB.Recordset
Dim RsSub As New ADODB.Recordset
Dim i As Double
On Error GoTo XErr
If Val(VsfNote.TextMatrix(VsfNote.Row, 4)) = 88 Or Val(VsfNote.TextMatrix(VsfNote.Row, 4)) = 1 Then
    With VsfDetail
        RsQ.Open "Select * From EntMst Where LCode=" & Val(VsfNote.TextMatrix(VsfNote.Row, 4)), DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        RsSub.Open "Select * From BSheetDtl Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQ.EOF = False
            .TextMatrix(.Row, 1) = RsQ.Fields("EName")
            If Val(VsfNote.TextMatrix(VsfNote.Row, 4)) = 1 Then
                If RsQ.Fields("ECode") = -465 And mProfit >= 0 Then
                    .TextMatrix(.Row, 0) = "Add"
                    .TextMatrix(.Row, 2) = mProfit
                ElseIf RsQ.Fields("ECode") = -466 And mProfit < 0 Then
                    .TextMatrix(.Row, 0) = "Less"
                    .TextMatrix(.Row, 2) = mProfit
                Else
                    If RsSub.BOF = False Then
                        RsSub.MoveFirst
                        RsSub.Find "ECode=" & RsQ.Fields("ECode")
                    End If
                    If RsSub.EOF = False Then
                        .TextMatrix(.Row, 0) = RsSub.Fields("Side")
                        .TextMatrix(.Row, 2) = RsSub.Fields("Amt")
                    Else
                        .TextMatrix(.Row, 2) = "0"
                    End If
                End If
            ElseIf Val(VsfNote.TextMatrix(VsfNote.Row, 4)) = 88 Then
                If RsQ.Fields("ECode") = -425 And mProfit >= 0 Then
                    .TextMatrix(.Row, 0) = "Add"
                    .TextMatrix(.Row, 2) = mProfit
                ElseIf RsQ.Fields("ECode") = -456 And mProfit < 0 Then
                    .TextMatrix(.Row, 0) = "Less"
                    .TextMatrix(.Row, 2) = mProfit
                Else
                    If RsSub.BOF = False Then
                        RsSub.MoveFirst
                        RsSub.Find "ECode=" & RsQ.Fields("ECode")
                    End If
                    If RsSub.EOF = False Then
                        .TextMatrix(.Row, 0) = RsSub.Fields("Side")
                        .TextMatrix(.Row, 2) = RsSub.Fields("Amt")
                    Else
                        .TextMatrix(.Row, 2) = "0"
                    End If
                End If
            Else
                If RsSub.BOF = False Then
                    RsSub.MoveFirst
                    RsSub.Find "ECode=" & RsQ.Fields("ECode")
                End If
                If RsSub.EOF = False Then
                    .TextMatrix(.Row, 0) = RsSub.Fields("Side")
                    .TextMatrix(.Row, 2) = RsSub.Fields("Amt")
                Else
                    .TextMatrix(.Row, 2) = "0"
                End If
            End If
            .TextMatrix(.Row, 3) = RsQ.Fields("ECode")
            .TextMatrix(.Row, 4) = "B"
            RsQ.MoveNext
            If RsQ.EOF = False Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        Loop
        If Val(VsfNote.TextMatrix(VsfNote.Row, 4)) = 88 Then
            i = .FindRow(IIf(Val(VsfNote.TextMatrix(VsfNote.Row, 2)) >= 0, -425, -456), 1, 3)
        Else
            i = .FindRow(IIf(Val(VsfNote.TextMatrix(VsfNote.Row, 2)) >= 0, -465, -466), 1, 3)
        End If
        Set RsQ = Nothing
        RsQ.Open "Select Side,Amt,ECode From JvDtl Where AcCode=" & mAcCode & " Order By ECode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQ.EOF = False
            i = .FindRow(RsQ.Fields("ECode"), 1, 3)
            If i > 0 Then
                If Val(.TextMatrix(i, 2)) = 0 Then
                    .TextMatrix(i, 0) = IIf(RsQ.Fields("Side") = "D", "Less", "Add")
                    .TextMatrix(i, 2) = RsQ.Fields("Amt")
                Else
                    If .TextMatrix(i, 0) = RsQ.Fields("Side") Then
                        .TextMatrix(i, 2) = Val(.TextMatrix(i, 2)) + RsQ.Fields("Amt")
                    Else
                        .TextMatrix(i, 2) = Val(.TextMatrix(i, 2)) - RsQ.Fields("Amt")
                        If Val(.TextMatrix(i, 2)) < 0 Then
                            If .TextMatrix(i, 0) = "Add" Then .TextMatrix(i, 0) = "Less" Else .TextMatrix(i, 0) = "Add"
                        End If
                    End If
                End If
            End If
            RsQ.MoveNext
        Loop
    End With
    SetDTotal
Else
    Set RsQ = Nothing
    RsQ.Open "Select IIF(Side=HSide,'Add','Less') As Side,EName,Amt,ECode From TmpCtDtl Where LCode=" & VsfNote.TextMatrix(VsfNote.Row, 4) & " Order By Side", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    With VsfDetail
        Do While RsQ.EOF = False
            .TextMatrix(.Row, 0) = RsQ.Fields("Side")
            .TextMatrix(.Row, 1) = RsQ.Fields("EName")
            .TextMatrix(.Row, 2) = RsQ.Fields("Amt")
            .TextMatrix(.Row, 3) = RsQ.Fields("ECode")
            .TextMatrix(.Row, 4) = "P"
            RsQ.MoveNext
            If RsQ.EOF = False Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        Loop
        Set RsQ = Nothing
        RsQ.Open "Select EName,ECode From EntMst Where LCode=" & VsfNote.TextMatrix(VsfNote.Row, 4) & " Order By EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQ.EOF = False
            i = .FindRow(RsQ.Fields("ECode"), 1, 3)
            If i < 0 Then
                If .TextMatrix(.Rows - 1, 1) <> "" Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                End If
                .TextMatrix(.Row, 1) = "Add"
                .TextMatrix(.Row, 1) = RsQ.Fields("EName")
                .TextMatrix(.Row, 3) = RsQ.Fields("ECode")
            End If
            RsQ.MoveNext
        Loop
    SetDTotal
    End With
End If
Exit Function
XErr:
MsgBox Err.Description
Resume
End Function

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
Private Sub CmdExport_Click()
    ExpMainData
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'BSHEET_IND','EXPORT_XLS','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
End Sub
Private Sub ExpMainData()
Dim i As Double
With VsfMainExport
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
    .Refresh
    .Rows = .Rows + 1
    .Row = 1
    For i = 1 To VsfHelp.Rows - 1
        .TextMatrix(i, 0) = VsfHelp.TextMatrix(i, 0)
        .TextMatrix(i, 1) = VsfHelp.TextMatrix(i, 1)
        .TextMatrix(i, 2) = VsfHelp.TextMatrix(i, 2)
        .TextMatrix(i, 3) = VsfHelp.TextMatrix(i, 3)
        .TextMatrix(i, 4) = VsfHelp.TextMatrix(i, 4)
        .TextMatrix(i, 5) = VsfHelp.TextMatrix(i, 5)
        .Rows = .Rows + 1
    Next
    .Rows = .Rows + 1
    .Refresh
    i = .Rows - 1
    Do While i > 0
        .TextMatrix(i, 0) = .TextMatrix(i - 1, 0)
        .TextMatrix(i, 1) = .TextMatrix(i - 1, 1)
        .TextMatrix(i, 2) = .TextMatrix(i - 1, 2)
        .TextMatrix(i, 3) = .TextMatrix(i - 1, 3)
        .TextMatrix(i, 4) = .TextMatrix(i - 1, 4)
        .TextMatrix(i, 5) = .TextMatrix(i - 1, 5)
        i = i - 1
    Loop
    .Cell(flexcpText, 0, 0, 0, .Cols - 1) = ""
    .TextMatrix(0, 0) = LsvClient.TextMatrix(LsvClient.Row, 0)
    .TextMatrix(0, 1) = LsvClient.TextMatrix(LsvClient.Row, 1)
End With
VsfMainExport.SaveGrid Environ("USERPROFILE") & "\Desktop\BSHEET.XLS", flexFileTabText
MsgBox "Successfully Excel File Generated In " + Environ("USERPROFILE") + "\Desktop\BSHEET.XLS", vbInformation, "Alert"
End Sub
Private Sub SetData1()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And SeqMst.HCode=" & _
"HedMst.HCode And HedMst.HType=1 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    RsQry.Open "Select HSide As AcSide,HCode As HeadCode,Sum(CBal) As TotRs From TmpTrialBal Where HType=1 Group By HSide,HCode Union All Select HSide As AcSide,HCode" & _
    " As HeadCode,Sum(DBal) As TotRs From TmpTrialBal Where HType=1 Group By HSide,HCode Order By HeadCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        If RsQry.Fields("AcSide") = "C" Then
            If RsQry.Fields("HeadCode") = 5 Or RsQry.Fields("HeadCode") = 13 Or RsQry.Fields("HeadCode") = 58 Then
                i = .FindRow(5, 1, 6)
            Else
                i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
            End If
            If i > 0 Then
                If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 13 Or RsQry.Fields("HeadCode") = 5 Then    '   Income And Expenditure
                    If mProfit = 0 Then
                        mProfit = SetProfit(mAcCode)
                        .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs")) + mProfit, "0.00")
                    Else
                        .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs")), "0.00")
                    End If
                Else
                    .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs")), "0.00")
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
    i = .FindRow(9, 1, 6)
    If i > 0 Then
        If Val(.TextMatrix(i, 2)) = 0 Then
            mProfit = SetProfit(mAcCode)
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
                    mProfit = SetProfit(mAcCode)
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
            If RsQry.Fields("HCode") = 57 Then
                i = .FindRow(13, 1, 6)
            ElseIf RsQry.Fields("HCode") = 58 Then
                i = .FindRow(5, 1, 6)
            Else
                i = .FindRow(RsQry.Fields("HCode"), 1, 6)
            End If
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
Private Sub PrintRec()
Dim i As Double
Dim mHead As String
Dim mClient1 As String
Dim RsRep As New ADODB.Recordset
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Balance Sheet does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
RsRep.Open "Select RepDtl.*,AdtMst.AdtName From RepDtl,AdtMst Where AcCode=" & mAcCode & " And RepDtl.AdtNo=AdtMst.AdtNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsRep.EOF = False Then
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        MsgBox "UDIN not issued. Non-UDIN Report will be printed.", vbInformation, "Information"
    End If
Else
    MsgBox "Signing Information not available. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
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
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\BSXRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    If Len(LsvClient.TextMatrix(LsvClient.Row, 6)) > 0 Then
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 6) & "'"
        .Formulas(8) = "mSubHead='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
        mClient1 = CStr(LsvClient.TextMatrix(LsvClient.Row, 6))
    ElseIf Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0 Then
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
        .Formulas(8) = "mSubHead=''"
        mClient1 = CStr(LsvClient.TextMatrix(LsvClient.Row, 3))
    Else
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
        .Formulas(8) = "mSubHead=''"
        mClient1 = CStr(LsvClient.TextMatrix(LsvClient.Row, 0)) & ", " & CStr(LsvClient.TextMatrix(LsvClient.Row, 2))
    End If
    .Formulas(9) = "mTitle1='Balance Sheet as on " & RsRep.Fields("RTDt") & "'"
    .Formulas(10) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(11) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
    .Formulas(12) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(13) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(14) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(15) = "mClient='" & IIf(Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0, TxtName.Text + IIf(Len(LsvClient.TextMatrix(LsvClient.Row, 2)) > 0, ", " + LsvClient.TextMatrix(LsvClient.Row, 2), ""), "") & "'"
    .Formulas(16) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(17) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(18) = "mLTitle='LIABILITIES'"
    .Formulas(19) = "mRTitle='ASSETS'"
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        If Len(LsvClient.TextMatrix(LsvClient.Row, 6)) > 0 Then
            mHead = LsvClient.TextMatrix(LsvClient.Row, 6)
        ElseIf Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0 Then
            mHead = LsvClient.TextMatrix(LsvClient.Row, 3)
        Else
            mHead = LsvClient.TextMatrix(LsvClient.Row, 0) & ", " & LsvClient.TextMatrix(LsvClient.Row, 2)
        End If
        .Formulas(20) = "mSub='This Balance Sheet is issued for the sole purpose of internal use of " & mHead & " and should not be presented before any third parties/agencies/authorities without our consent.'"
    ElseIf LsvClient.TextMatrix(LsvClient.Row, 4) = "FC" Then
        .Formulas(20) = "mSub='The above Balance Sheet, to the best of our belief, contains a true account of the Funds and Liabilities and the Property and Assets of the Foreign Contribution Account of the Trust.'"
    ElseIf LsvClient.TextMatrix(LsvClient.Row, 4) = "Library" Then
        .Formulas(20) = "mSub='The above Balance Sheet, to the best of our belief, contains a true account of the Funds and Liabilities and the Property and Assets of the library.'"
    Else: .Formulas(20) = "mSub='The above Balance Sheet, to the best of our belief, contains a true account of the Funds and Liabilities and the Property and Assets of the branch.'"
    End If
        If LsvClient.TextMatrix(LsvClient.Row, 4) = "Library" Then
        .Formulas(21) = "mTSign='Trustee/Secretary'"
    ElseIf LsvClient.TextMatrix(LsvClient.Row, 4) = "FC" Then
        .Formulas(21) = "mTSign='Chief Functionary/Trustee'"
    ElseIf LsvClient.TextMatrix(LsvClient.Row, 4) = "School" Then
        .Formulas(21) = "mTSign='Principal/In-charge'"
    Else: .Formulas(21) = "mTSign='Trustee/In-charge'"
    End If
    If LsvClient.TextMatrix(LsvClient.Row, 4) = "School" Then .Formulas(22) = "mClient1='" & TxtName.Text + IIf(Len(LsvClient.TextMatrix(LsvClient.Row, 2)) > 0, ", " + LsvClient.TextMatrix(LsvClient.Row, 2), "") & "'" Else .Formulas(22) = "mClient1='" & CStr(mClient1) & "'"
    .Action = 1
    For i = 0 To 22
        .Formulas(i) = ""
    Next
If (LsvClient.TextMatrix(LsvClient.Row, 4) = "Trust" Or LsvClient.TextMatrix(LsvClient.Row, 4) = "FC") Then NotePrint1 Else NotePrint
End With
End Sub
Private Sub NotePrint()
Dim i As Integer
Dim mGroup As Double
Dim mTotal As Double
Dim RsDataO As New ADODB.Recordset
Dim RsDataD As New ADODB.Recordset
Dim RsQData As New ADODB.Recordset
Dim mRAmt As Double
Dim mSAmt As Double
Dim mCAmt As Double
Dim mDAmt As Double
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpNotePrn"
With VsfHelp
    If Val(.TextMatrix(1, 1)) <> 0 Then
        i = .FindRow(14, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,OpDtl.OpBal As RTotal From OpDtl,LedMst Where OpDtl.AcCode=" & mAcCode & " And LedMst.LCode=OpDtl.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','  Opening Balance'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            RsDataO.Open "Select IIF(TmpCtDtl.Side=TmpCtDtl.HSide,'Add','Less') As TrnType,TmpCtDtl.EName,TmpCtDtl.Amt As Amount,TmpCtDtl.ECode,LedMst.LName From TmpCtDtl,LedMst Where " & _
            "TmpCtDtl.LCode=LedMst.LCode And TmpCtDtl.HCode=" & mGroup & " Order By TmpCtDtl.Side", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' " & RsDataO.Fields("TrnType") & _
                    ":" & Space(1) & RsDataO.Fields("EName") & "'," & IIf(RsDataO.Fields("TrnType") = "Add", RsDataO.Fields("Amount"), RsDataO.Fields("Amount") * -1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,RName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "'," & RsDataO.Fields("RTotal") & _
                    "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'Closing Balance')"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(2, 1)) <> 0 Then
        i = .FindRow(2, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsDataO.Fields("LName") & "'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsDataO.Fields("LName") & "'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsDataO.Fields("LName") & "'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(5, 1)) <> 0 Then
        i = .FindRow(5, 1, 6)
        If i = -1 Then i = .FindRow(58, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,OpDtl.OpBal As RTotal From OpDtl,LedMst Where OpDtl.AcCode=" & mAcCode & " And LedMst.LCode=OpDtl.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsDataO.EOF = False Then
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' Opening Balance'," & RsDataO.Fields("RTotal") & _
                    "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
            End If
            Set RsDataO = Nothing
            RsDataO.Open "Select IIF(TmpCtDtl.Side=TmpCtDtl.HSide,'Add','Less') As TrnType,TmpCtDtl.EName,IIF(TmpCtDtl.Side=TmpCtDtl.HSide,TmpCtDtl.Amt,TmpCtDtl.Amt*-1) As Amount,TmpCtDtl.ECode,LedMst.LName From TmpCtDtl," & _
            "LedMst Where TmpCtDtl.LCode=LedMst.LCode And TmpCtDtl.HCode=" & mGroup & " Order By TmpCtDtl.Side", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsDataO.Fields("TrnType") & _
                    ":" & Space(1) & RsDataO.Fields("EName") & "'," & RsDataO.Fields("Amount") & "," & IIf(RsDataO.Fields("TrnType") = "Add", 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            mTotal = SetProfit(mAcCode)
            DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & IIf(Round(mTotal, 2) >= 0, "Add: Surplus brought from Income and Expenditure Account", "Less: Deficit brought from Income and Expenditure Account") & _
            "'," & mTotal & "," & IIf(Round(mTotal, 2) >= 0, 0, 1) & ")"
        End If
    End If
    If Val(.TextMatrix(1, 4)) <> 0 Then
        i = .FindRow(6, 1, 7)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 7))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.OpDr As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=TmpTrialBal.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mRAmt = RsQData.Fields("RTotal")
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From TmpCtDtl Where LCode=" & RsDataO.Fields("LCode") & " And EName Like '*Sale*'", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mSAmt = 0
                Do While RsQData.EOF = False
                    mSAmt = mSAmt + RsQData.Fields("RTotal")
                    RsQData.MoveNext
                Loop
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mDAmt = RsQData.Fields("RTotal") - mSAmt
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(DBal)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mCAmt = RsQData.Fields("RTotal")
                DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,LName,SAmt,CAmt,DAmt,HCode) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                "'," & RsDataO.Fields("RTotal") & "," & mRAmt & ",'" & RsDataO.Fields("LName") & "'," & mSAmt & "," & mCAmt & "," & mDAmt & "," & mGroup & ")"
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(2, 4)) <> 0 Then
        i = .FindRow(7, 1, 7)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 7))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.OpDr As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mRAmt = RsQData.Fields("RTotal")
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From TmpCtDtl Where LCode=" & RsDataO.Fields("LCode") & " And EName Like '*Sale*'", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mSAmt = 0
                Do While RsQData.EOF = False
                    mSAmt = mSAmt + RsQData.Fields("RTotal")
                    RsQData.MoveNext
                Loop
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mDAmt = RsQData.Fields("RTotal") - mSAmt
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mCAmt = RsQData.Fields("RTotal")
                DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,LName,SAmt,CAmt,DAmt,HCode) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                "'," & RsDataO.Fields("RTotal") & "," & mRAmt & ",'" & RsDataO.Fields("LName") & "'," & mSAmt & "," & mCAmt & "," & mDAmt & "," & mGroup & ")"
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(4, 4)) <> 0 Then
        i = .FindRow(9, 4, 7)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 7))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.LCode In (10000,10001,10315) Order By TmpTrialBal.LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("DBal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "',' " & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("DBal") & ")"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.HCode=" & mGroup & " And TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.LCode" & _
            " Not In (10000,10001,10315) Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("DBal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "',' " & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("DBal") & ")"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
End With
DbDataDB.CommitTrans
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
    .Formulas(12) = "mTrack='0'"
    .Formulas(13) = "mType='B'"
    .Formulas(14) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Action = 1
    For i = 0 To 14
        .Formulas(i) = ""
    Next
End With
End Sub
Private Sub NotePrint1()
Dim i As Integer
Dim mGroup As Double
Dim mTotal As Double
Dim RsDataO As New ADODB.Recordset
Dim RsDataD As New ADODB.Recordset
Dim RsQData As New ADODB.Recordset
Dim mRAmt As Double
Dim mSAmt As Double
Dim mCAmt As Double
Dim mDAmt As Double
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpNotePrn"
With VsfHelp
    If Val(.TextMatrix(1, 1)) <> 0 Then
        i = .FindRow(1, 1, 6)
        If mFinYear = "19-20" Then
            If i = -1 Then i = .FindRow(57, 1, 6)
        Else
            If i = -1 Then i = .FindRow(13, 1, 6)
        End If
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,OpDtl.OpBal As RTotal From OpDtl,LedMst Where OpDtl.AcCode=" & mAcCode & " And LedMst.LCode=OpDtl.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','  Opening Balance'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            RsDataO.Open "Select IIF(TmpCtDtl.Side=TmpCtDtl.HSide,'Add','Less') As TrnType,TmpCtDtl.EName,TmpCtDtl.Amt As Amount,TmpCtDtl.ECode,LedMst.LName From TmpCtDtl,LedMst Where " & _
            "TmpCtDtl.LCode=LedMst.LCode And TmpCtDtl.HCode=" & mGroup & " Order By TmpCtDtl.Side", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' " & RsDataO.Fields("TrnType") & _
                    ":" & Space(1) & RsDataO.Fields("EName") & "'," & IIf(RsDataO.Fields("TrnType") = "Add", RsDataO.Fields("Amount"), RsDataO.Fields("Amount") * -1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,RName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "'," & RsDataO.Fields("RTotal") & _
                    "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'Closing Balance')"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(2, 1)) <> 0 Then
        i = .FindRow(14, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,OpDtl.OpBal As RTotal From OpDtl,LedMst Where OpDtl.AcCode=" & mAcCode & " And LedMst.LCode=OpDtl.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','  Opening Balance'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            RsDataO.Open "Select IIF(TmpCtDtl.Side=TmpCtDtl.HSide,'Add','Less') As TrnType,TmpCtDtl.EName,TmpCtDtl.Amt As Amount,TmpCtDtl.ECode,LedMst.LName From TmpCtDtl,LedMst Where " & _
            "TmpCtDtl.LCode=LedMst.LCode And TmpCtDtl.HCode=" & mGroup & " Order By TmpCtDtl.Side", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' " & RsDataO.Fields("TrnType") & _
                    ":" & Space(1) & RsDataO.Fields("EName") & "'," & IIf(RsDataO.Fields("TrnType") = "Add", RsDataO.Fields("Amount"), RsDataO.Fields("Amount") * -1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,RName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "'," & RsDataO.Fields("RTotal") & _
                    "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'Closing Balance')"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(3, 1)) <> 0 Then
        i = .FindRow(2, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsDataO.Fields("LName") & "'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(4, 1)) <> 0 Then
        i = .FindRow(3, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsDataO.Fields("LName") & "'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(5, 1)) <> 0 Then
        i = .FindRow(4, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsDataO.Fields("LName") & "'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(6, 1)) <> 0 Then
        i = .FindRow(5, 1, 6)
        If i = -1 Then i = .FindRow(58, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,OpDtl.OpBal As RTotal From OpDtl,LedMst Where OpDtl.AcCode=" & mAcCode & " And LedMst.LCode=OpDtl.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsDataO.EOF = False Then
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' Opening Balance'," & RsDataO.Fields("RTotal") & _
                    "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
            End If
            Set RsDataO = Nothing
            RsDataO.Open "Select IIF(TmpCtDtl.Side=TmpCtDtl.HSide,'Add','Less') As TrnType,TmpCtDtl.EName,IIF(TmpCtDtl.Side=TmpCtDtl.HSide,TmpCtDtl.Amt,TmpCtDtl.Amt*-1) As Amount,TmpCtDtl.ECode,LedMst.LName From TmpCtDtl," & _
            "LedMst Where TmpCtDtl.LCode=LedMst.LCode And TmpCtDtl.HCode=" & mGroup & " Order By TmpCtDtl.Side", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsDataO.Fields("TrnType") & _
                    ":" & Space(1) & RsDataO.Fields("EName") & "'," & RsDataO.Fields("Amount") & "," & IIf(RsDataO.Fields("TrnType") = "Add", 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            mTotal = SetProfit(mAcCode)
            DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & IIf(Round(mTotal, 2) >= 0, "Add: Surplus brought from Income and Expenditure Account", "Less: Deficit brought from Income and Expenditure Account") & _
            "'," & mTotal & "," & IIf(Round(mTotal, 2) >= 0, 0, 1) & ")"
        End If
    End If
    If Val(.TextMatrix(1, 4)) <> 0 Then
        i = .FindRow(6, 1, 7)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 7))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.OpDr As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=TmpTrialBal.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mRAmt = RsQData.Fields("RTotal")
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From TmpCtDtl Where LCode=" & RsDataO.Fields("LCode") & " And EName Like '*Sale*'", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mSAmt = 0
                Do While RsQData.EOF = False
                    mSAmt = mSAmt + RsQData.Fields("RTotal")
                    RsQData.MoveNext
                Loop
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mDAmt = RsQData.Fields("RTotal") - mSAmt
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(DBal)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mCAmt = RsQData.Fields("RTotal")
                DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,LName,SAmt,CAmt,DAmt,HCode) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                "'," & RsDataO.Fields("RTotal") & "," & mRAmt & ",'" & RsDataO.Fields("LName") & "'," & mSAmt & "," & mCAmt & "," & mDAmt & "," & mGroup & ")"
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(2, 4)) <> 0 Then
        i = .FindRow(7, 1, 7)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 7))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.OpDr As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And LedMst.HCode=" & _
            mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mRAmt = RsQData.Fields("RTotal")
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From TmpCtDtl Where LCode=" & RsDataO.Fields("LCode") & " And EName Like '*Sale*'", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mSAmt = 0
                Do While RsQData.EOF = False
                    mSAmt = mSAmt + RsQData.Fields("RTotal")
                    RsQData.MoveNext
                Loop
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mDAmt = RsQData.Fields("RTotal") - mSAmt
                Set RsQData = Nothing
                RsQData.Open "Select IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) AS RTotal From TmpTrialBal Where LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                mCAmt = RsQData.Fields("RTotal")
                DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,LName,SAmt,CAmt,DAmt,HCode) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                "'," & RsDataO.Fields("RTotal") & "," & mRAmt & ",'" & RsDataO.Fields("LName") & "'," & mSAmt & "," & mCAmt & "," & mDAmt & "," & mGroup & ")"
                RsDataO.MoveNext
            Loop
        End If
    End If
    If Val(.TextMatrix(4, 4)) <> 0 Then
        i = .FindRow(9, 4, 7)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 7))
            Set RsDataO = Nothing
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & _
            " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
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
            RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.LCode In (10000,10001,10315) Order By TmpTrialBal.LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("DBal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "',' " & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("DBal") & ")"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.HCode=" & mGroup & " And TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.LCode" & _
            " Not In (10000,10001,10315) Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsDataO.EOF = False
                If RsDataO.Fields("DBal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "',' " & RsDataO.Fields("LName") & _
                    "'," & RsDataO.Fields("DBal") & ")"
                End If
                RsDataO.MoveNext
            Loop
        End If
    End If
End With
DbDataDB.CommitTrans
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
    .Formulas(12) = "mTrack='0'"
    .Formulas(13) = "mType='B'"
    .Formulas(14) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Action = 1
    For i = 0 To 14
        .Formulas(i) = ""
    Next
End With
End Sub


