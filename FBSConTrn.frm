VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmBSCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sheet [Consolidated]"
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
   Icon            =   "FBSConTrn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   15690
   Begin VB.Frame FraNote 
      Height          =   6612
      Left            =   18000
      TabIndex        =   10
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         Picture         =   "FBSConTrn.frx":000C
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
         FormatString    =   $"FBSConTrn.frx":044E
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
      TabIndex        =   8
      Top             =   1920
      Width           =   14750
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   3852
         Left            =   120
         TabIndex        =   9
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
         FormatString    =   $"FBSConTrn.frx":0497
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
      TabIndex        =   6
      Top             =   0
      Width           =   15492
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame FraList 
         Height          =   3972
         Left            =   24000
         TabIndex        =   26
         Top             =   1200
         Width           =   7572
         Begin VB.TextBox TxtLTotal 
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
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   3500
            Width           =   1776
         End
         Begin VB.CommandButton CmdLClose 
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
            Left            =   1608
            Picture         =   "FBSConTrn.frx":04E0
            TabIndex        =   27
            ToolTipText     =   "Cancel"
            Top             =   3500
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfList 
            Height          =   3252
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   7300
            _cx             =   12876
            _cy             =   5736
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
            BackColor       =   15007437
            ForeColor       =   -2147483640
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   128
            BackColorBkg    =   15007437
            BackColorAlternate=   15007437
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
            FormatString    =   $"FBSConTrn.frx":0922
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
            Height          =   3852
            Index           =   4
            Left            =   0
            Top             =   120
            Width           =   7572
         End
      End
      Begin VB.Frame FraDetail 
         Height          =   6132
         Left            =   18000
         TabIndex        =   14
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
            TabIndex        =   17
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
            Picture         =   "FBSConTrn.frx":096B
            TabIndex        =   16
            ToolTipText     =   "Cancel"
            Top             =   5640
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfDetail 
            Height          =   5412
            Left            =   120
            TabIndex        =   15
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
            FormatString    =   $"FBSConTrn.frx":0DAD
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
      Begin VB.CommandButton CmdSave 
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   732
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
         TabIndex        =   24
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   1470
      End
      Begin VB.Frame FraMainExp 
         BackColor       =   &H00F7E0FE&
         Height          =   5412
         Left            =   1.80000e5
         TabIndex        =   20
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
            TabIndex        =   21
            ToolTipText     =   "Cancel"
            Top             =   4920
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfMainExport 
            Height          =   5052
            Left            =   0
            TabIndex        =   22
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
            FormatString    =   $"FBSConTrn.frx":0DF6
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
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   810
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
         Picture         =   "FBSConTrn.frx":0E3F
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
         TabIndex        =   5
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
         FormatString    =   $"FBSConTrn.frx":1475
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
         TabIndex        =   23
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
         TabIndex        =   7
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
Attribute VB_Name = "FrmBSCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim mAcCode As Double
Dim mAcList As String
Dim RsLedger As New ADODB.Recordset
Dim RsClient As New ADODB.Recordset
Dim RsState As New ADODB.Recordset
Dim mProfit As Double
Dim RsRep As New ADODB.Recordset
Private Sub CmdCancelD_Click()
    FraDetail.Left = 18000
    FraNote.Left = 180
    VsfNote.SetFocus
End Sub

Private Sub CmdLClose_Click()
    FraList.Left = 36000
End Sub

Private Sub CmdSearch_Click()
    FraClientHelp.Left = 180
    LsvClient.SetFocus
End Sub
Private Sub Form_Load()
    DBWorkTmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & App.Path + "\LocalDB.Mdb'"
    DBWorkTmp.Open
    Me.Left = 50
    Me.Top = 50
    SetCombo
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 " & _
"And AcMst.AcType=1 And AcMst.AcType=GrpMst.GCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.City,AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,'')," & _
"GrpMst.GName,AcMst.AcCode From AcMst,GrpMst,AcMst As AcMst1 Where AcMst.PaCode=AcMst1.AcCode And AcMst.AcType=2 And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    FraClientHelp.Left = 18000
    SetParent
    If RsClient.BOF = False Then
        RsClient.MoveFirst
        RsClient.Find "Accode=" & mAcCode
    End If
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
    If LsvClient.TextMatrix(LsvClient.Row, 4) <> "Trust" Then SetData1 Else SetData
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'BSHEET_CON','VIEW','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    VsfHelp.SetFocus
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Cancel"
        FraNote.Left = 18000
        FraList.Left = 3000
    Case "Print"
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'BSHEET_CON','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
    Case "Exit"
        DBWorkTmp.Close
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(2).Enabled = True
End Function
Private Sub SetData()
Dim RsExclude As New ADODB.Recordset
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & _
mAcCode & ") And SeqMst.HCode=HedMst.HCode And HedMst.HType=1 And HedMst.HCode<>14 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    RsQry.Open "Select HSide As AcSide,HCode As HeadCode,Sum(CBal) As TotRs From TmpTrialBal Where HType=1 And HCode<>14 Group By HSide,HCode Union All Select HSide,HCode," & _
    "Sum(DBal) From TmpTrialBal Where HType=1 And HCode<>14 Group By HSide,HCode Order By HeadCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
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
                    If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 57 Or RsQry.Fields("HeadCode") = 13 Or RsQry.Fields("HeadCode") = 5 Then    '   Income And Expenditure
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
Private Sub SetData1()
Dim RsQry As New ADODB.Recordset
If mFinYear = "19-20" Then
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode=6 And SeqMst.HCode=HedMst.HCode And HedMst.HType=1 And HedMst.HCode<>14 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Else
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & _
    mAcCode & ") And SeqMst.HCode=HedMst.HCode And HedMst.HType=1 And HedMst.HCode<>14 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
End If
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
                If RsQry.Fields("HeadCode") = 5 Or RsQry.Fields("HeadCode") = 13 Or RsQry.Fields("HeadCode") = 58 Then
                    i = .FindRow(58, 1, 6)
                Else
                    i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
                End If
            Else
                i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
            End If
0            If i > 0 Then
                If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 13 Or RsQry.Fields("HeadCode") = 5 Then    '   Income And Expenditure
                    If mProfit = 0 Then
                        mProfit = SetProfitAll(mAcList)
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
If Val(TxtCTotal.Text) - Val(TxtDTotal.Text) = 0 Then LblTotal.Caption = "" Else LblTotal.Caption = "Total mismatching Of Rs.-->" & Format(CStr(Abs(Val(TxtCTotal.Text) - Val(TxtDTotal.Text))), "0.00")
End Sub
Private Sub CmdExport_Click()
    ExpMainData
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'BSHEET_CON','EXPORT_XLS','" & Date & "','" & Time & "')"
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
Private Sub SetParent()
Dim RsQ As New ADODB.Recordset
If LsvClient.TextMatrix(LsvClient.Row, 4) <> "Trust" Then
    mAcList = mAcCode
    RsQ.Open "Select * From QGroup Where PACode=" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Else
    mAcList = ""
    RsQ.Open "Select * From QGroup Where SACode=" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
End If
Do While RsQ.EOF = False
    If Len(mAcList) = 0 Then
        mAcList = CStr(RsQ.Fields("AcCode"))
    Else
        mAcList = mAcList + "," + CStr(RsQ.Fields("AcCode"))
    End If
    RsQ.MoveNext
Loop
Set RsQ = Nothing
RsQ.Open "Select FileNo,AcName,0 As Amount,AcCode From AcMst Where AcCode In (" & mAcList & ") Order By FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfList.DataSource = RsQ
With VsfList
    .TextMatrix(0, 0) = "FILE NO"
    .ColWidth(0) = 1200
    .TextMatrix(0, 1) = "NAME"
    .ColWidth(1) = 4500
    .TextMatrix(0, 2) = "AMOUNT"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColWidth(3) = 0   'ACCODE
    .Refresh
End With
End Sub
Private Sub CmdSave_Click()
Dim i As Integer
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete NtDtl.* From NtDtl,HedMst Where NtDtl.RType=1 And NtDtl.AcCode=" & mAcCode & " And NtDtl.HCode=HedMst.HCode And HedMst.HType=1"
With VsfHelp
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 Then
            DbDataDB.Execute "Insert InTo NtDtl (AcCode,HCode,[Note],RSide,RType) Values (" & mAcCode & "," & Val(.TextMatrix(i, 6)) & "," & Val(.TextMatrix(i, 1)) & ",'C',1)"
        End If
    Next
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 4)) <> 0 Then
            DbDataDB.Execute "Insert InTo NtDtl (AcCode,HCode,[Note],RSide,RType) Values (" & mAcCode & "," & Val(.TextMatrix(i, 7)) & "," & Val(.TextMatrix(i, 4)) & ",'D',1)"
        End If
    Next
End With
'DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'BSHEET_CON','SAVE_NOTES','" & Date & "','" & Time & "')"
DbDataDB.CommitTrans
MsgBox "Record Saved Successfully.", vbInformation, "Alert"
Exit Sub
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Sub

Private Sub VsfHelp_DblClick()
    FraList.Left = 3000
    SetCmpTotal
    VsfList.SetFocus
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
Private Sub VsfList_DblClick()
    FraList.Left = 36000
    FraNote.Left = 180
    SetLedger VsfHelp.Col
    If VsfNote.TextMatrix(1, 0) <> "" Then
        FraNote.Left = 180
        VsfNote.SetFocus
    End If
End Sub
Private Sub SetLedger(ByVal mSide As Integer)
Dim RsQry As New ADODB.Recordset
Dim mParentcd As Double
Dim mLCode As Double
Dim mRow As Double
Dim mHeadCode As Double
If InStr(1, mAcList, ",") > 0 Then mParentcd = Mid(mAcList, 1, InStr(1, mAcList, ",") - 1) Else mParentcd = mAcList
mLCode = Val(VsfList.TextMatrix(VsfList.Row, 3))
If mSide <= 2 Then mHeadCode = Val(VsfHelp.TextMatrix(VsfHelp.Row, 6)) Else mHeadCode = Val(VsfHelp.TextMatrix(VsfHelp.Row, 7))
Set RsLedger = Nothing
If mHeadCode = 5 Then
    If mYear = "19-20" Then
        RsQry.Open "Select * From AcMst Where AcCode=" & mParentcd, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQry.Fields("AcType") = 2 Then
            RsLedger.Open "Select LedMst.LName,LedMst.LCode,HedMst.HCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 " & _
            "And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mLCode & ") And LedMst.HCode In (58,13,5) And " & _
            "LedMst.HCode=SeqMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mLCode & _
            ") And SeqMst.HCode=HedMst.HCode Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        End If
    Else
        RsLedger.Open "Select LedMst.LName,LedMst.LCode,HedMst.HCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 " & _
        "And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mLCode & ") And LedMst.HCode In (58,57,5) And " & _
        "LedMst.HCode=SeqMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mLCode & _
        ") And SeqMst.HCode=HedMst.HCode Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    End If
ElseIf mHeadCode = 1 Then
    RsLedger.Open "Select LedMst.LName,LedMst.LCode,HedMst.HCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 " & _
    "And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mLCode & ") And LedMst.HCode In (13,1) And " & _
    "LedMst.HCode=SeqMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mLCode & _
    ") And SeqMst.HCode=HedMst.HCode Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Else
    RsLedger.Open "Select LedMst.LName,LedMst.LCode,HedMst.HCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 " & _
    "And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mLCode & ") And LedMst.HCode=" & mHeadCode & _
    " And LedMst.HCode=SeqMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mLCode & _
    ") And SeqMst.HCode=HedMst.HCode Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
End If
Set RsQry = Nothing
With VsfNote
    .Editable = flexEDNone
    .Cols = 6
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
    .ColWidth(5) = 0   '  HCODE
    .Col = 0
    .Refresh
    .Rows = 2
    .Row = 1
    Do While RsLedger.EOF = False
        .TextMatrix(.Row, 0) = RsLedger.Fields("LName")
        .TextMatrix(.Row, 4) = RsLedger.Fields("LCode")
        .TextMatrix(.Row, 5) = RsLedger.Fields("HCode")
        RsLedger.MoveNext
        If RsLedger.EOF = False Then
            .Rows = .Rows + 1
            .Row = .Row + 1
        End If
    Loop
    For mHeadCode = 1 To .Rows - 1
        Set RsLedger = Nothing
        RsLedger.Open "Select * From TmpTrialBal Where AcCode=" & mLCode & " And LCode=" & Val(.TextMatrix(mHeadCode, 4)), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsLedger.EOF = False Then
            If mSide > 2 Then
                If Val(VsfHelp.TextMatrix(VsfHelp.Row, 7)) = 12 Then
                    .TextMatrix(mHeadCode, 1) = RsLedger.Fields("OpDr")
                    .TextMatrix(mHeadCode, 3) = RsLedger.Fields("DBal")
                Else
                    .TextMatrix(mHeadCode, 1) = RsLedger.Fields("OpDr")
                    .TextMatrix(mHeadCode, 2) = RsLedger.Fields("ADr") - RsLedger.Fields("ACr")
                    .TextMatrix(mHeadCode, 3) = RsLedger.Fields("DBal")
                End If
            Else
                .TextMatrix(mHeadCode, 1) = RsLedger.Fields("OpCr")
                .TextMatrix(mHeadCode, 2) = RsLedger.Fields("ACr") - RsLedger.Fields("ADr")
                .TextMatrix(mHeadCode, 3) = RsLedger.Fields("CBal")
            End If
        End If
        If Val(.TextMatrix(mHeadCode, 4)) = 1 Or Val(.TextMatrix(mHeadCode, 4)) = 88 Then
        .TextMatrix(mHeadCode, 2) = Val(.TextMatrix(mHeadCode, 2)) + SetProfitAll(mAcCode)
        .TextMatrix(mHeadCode, 3) = Val(.TextMatrix(mHeadCode, 3)) + SetProfitAll(mAcCode)
        End If
    Next
End With
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
        TxtNTotal.Text = Val(TxtNTotal.Text) + Val(.TextMatrix(i, 3))
    Next
End With
TxtOTotal.Text = Format(TxtOTotal.Text, "0.00")
TxtTotal.Text = Format(TxtTotal.Text, "0.00")
TxtNTotal.Text = Format(TxtNTotal.Text, "0.00")
End Sub
Private Sub SetCmpTotal()
Dim RsTotal As New ADODB.Recordset
Dim RsQry As New ADODB.Recordset
Dim RsExclude As New ADODB.Recordset
Dim mParentcd As Double
Dim mAmt As Double
Dim i As Integer
Dim m As Integer
Dim mGrp As Double
If InStr(1, mAcList, ",") > 0 Then mParentcd = Mid(mAcList, 1, InStr(1, mAcList, ",") - 1) Else mParentcd = mAcList
If VsfHelp.Col <= 2 Then
    mGrp = Val(VsfHelp.TextMatrix(VsfHelp.Row, 6))
    RsTotal.Open "Select AcCode,IIF(IsNull(Sum(CBal))=False,Sum(CBal),0) As RTotal From TmpTrialBal Where HCode=" & mGrp & " Group By AcCode Order By AcCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
Else
    mGrp = Val(VsfHelp.TextMatrix(VsfHelp.Row, 7))
    RsTotal.Open "Select AcCode,IIF(IsNull(Sum(DBal))=False,Sum(DBal),0) As RTotal From TmpTrialBal Where HCode=" & mGrp & " Group By AcCode Order By AcCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
End If
With VsfList
    For i = 1 To .Rows - 1
        .TextMatrix(i, 2) = "0.00"
    Next
    Do While RsTotal.EOF = False
        i = .FindRow(RsTotal.Fields("AcCode"), 1, 3)
        If i > 0 Then
            .TextMatrix(i, 2) = Round(RsTotal.Fields("RTotal"), 2)
        End If
        RsTotal.MoveNext
    Loop
    If mGrp = 5 Then
        For i = 1 To .Rows - 1
            mAcCode = Val(.TextMatrix(i, 3))
            mAmt = SetProfitAll(mAcCode)
            Set RsQry = Nothing
            RsQry.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And HCode In (57,58)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            .TextMatrix(m, 2) = Round(Val(.TextMatrix(m, 2)) + RsQry.Fields("RTotal"), 2)
            m = .FindRow(Val(.TextMatrix(i, 3)), 1, 3)
            If m > 0 Then
                .TextMatrix(m, 2) = Round(Val(.TextMatrix(m, 2)) + mAmt, 2)
                Set RsExclude = Nothing
                If mYear = "19-20" Then
                    Set RsQry = Nothing
                    RsQry.Open "Select * From AcMst Where AcCode=" & mParentcd, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                    If RsQry.Fields("AcType") = 2 Then
                        RsExclude.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode=" & mAcCode & " And HCode=57", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                        If RsExclude.EOF = False Then .TextMatrix(m, 2) = Round(Val(.TextMatrix(m, 2)) - RsExclude.Fields("TotRs"), 2)
                        Set RsExclude = Nothing
                        RsExclude.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode=" & mAcCode & " And HCode=13", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                        If RsExclude.EOF = False Then .TextMatrix(m, 2) = Round(Val(.TextMatrix(m, 2)) + RsExclude.Fields("TotRs"), 2)
                    End If
                End If
            End If
        Next
    End If
    If mGrp = 1 Then
        For i = 1 To .Rows - 1
            mAcCode = Val(.TextMatrix(i, 3))
            Set RsExclude = Nothing
            RsExclude.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode=" & mAcCode & " And HCode=13", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            i = .FindRow(mAcCode, 1, 3)
            If i > 0 And RsExclude.EOF = False Then .TextMatrix(i, 2) = Round(Val(.TextMatrix(i, 2)) + RsExclude.Fields("TotRs"), 2)
        Next
    End If
    TxtLTotal.Text = "0.00"
    For i = 1 To .Rows - 1
        TxtLTotal.Text = Val(Format(TxtLTotal.Text)) + Val(.TextMatrix(i, 2))
    Next
    TxtLTotal.Text = Format(TxtLTotal.Text, "0.00")
End With
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
Private Function ShowDetail()
Dim RsQ As New ADODB.Recordset
Dim RsSub As New ADODB.Recordset
Dim i As Double
Dim mLCode As Double
mLCode = Val(VsfList.TextMatrix(VsfList.Row, 3))
On Error GoTo XErr
Set RsQ = Nothing
RsQ.Open "Select IIF(Side=HSide,'Add','Less') As Side,EName,Amt,ECode From TmpCtDtl Where AcCode=" & mLCode & " And LCode=" & VsfNote.TextMatrix(VsfNote.Row, 4) & " Order By Side", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With VsfDetail
    Do While RsQ.EOF = False
        .TextMatrix(.Row, 0) = RsQ.Fields("Side")
        .TextMatrix(.Row, 1) = RsQ.Fields("EName")
        .TextMatrix(.Row, 2) = RsQ.Fields("Amt")
        .TextMatrix(.Row, 3) = RsQ.Fields("ECode")
        RsQ.MoveNext
        If RsQ.EOF = False Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
    Loop
End With
If Val(VsfNote.TextMatrix(VsfNote.Row, 4)) = 88 Or Val(VsfNote.TextMatrix(VsfNote.Row, 4)) = 1 Then
    With VsfDetail
        .Rows = .Rows + 1
        .Row = .Rows - 1
        mProfit = SetProfitAll(mAcCode)
        If mProfit > 0 Then
            .TextMatrix(.Row, 0) = "Add"
            .TextMatrix(.Row, 1) = "Surplus brought from Income and Expenditure Account"
            .TextMatrix(.Row, 2) = mProfit
        Else
            .TextMatrix(.Row, 0) = "Less"
            .TextMatrix(.Row, 1) = "Deficit brought from Income and Expenditure Account"
            .TextMatrix(.Row, 2) = mProfit
        End If
    End With
End If
SetDTotal
Exit Function
XErr:
MsgBox Err.Description
End Function
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
Private Sub PrintRec()
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Balance Sheet does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
Set RsRep = Nothing
RsRep.Open "Select RepDtl.*,AdtMst.AdtName From RepDtl,AdtMst Where AcCode=" & mAcCode & " And RepDtl.AdtNo=AdtMst.AdtNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsRep.EOF = False Then
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        MsgBox "UDIN not issued. Report Cannot be printed.", vbInformation, "Information"
        Exit Sub
    End If
Else
    MsgBox "Signing Information not available. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
Set RsState = Nothing
RsState.Open "Select * from SActMst Where State='" & RsClient.Fields("State") & "'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
BSheetPrint
NotePrint
End Sub
Private Sub NotePrint()
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
DbDataDB.Execute "Delete From TmpNotePrn"
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
With VsfHelp
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
                mAcName = Space(3) + LsvClient.TextMatrix(LsvClient.Row, 0)
            ElseIf RsClient.Fields("AcName") = "Trust Account" Then
                mAcName = Space(2) + RsClient.Fields("AcName")
            ElseIf RsClient.Fields("AcName") = "Foreign Contribution Account" Then
                mAcName = Space(1) + RsClient.Fields("AcName")
            Else
                mAcName = RsClient.Fields("AcName") + IIf(Len(RsClient.Fields("City")) <> 0, (", " + RsClient.Fields("City")), "")
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
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
                    RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
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
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
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
                    RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From TmpCtDtl Where TmpCtDtl.AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode") & _
                    " And EName Like '*Sale*'", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
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
                    RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mRAmt = 0
                    mRAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From TmpCtDtl Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode") & _
                    " And EName Like '*Sale*'", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
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
        If Val(.TextMatrix(4, 4)) <> 0 Then
            i = .FindRow(9, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,LedMst.HCode,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & _
                " And LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode<>14 And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
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
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & _
                        .TextMatrix(i, 3) & "','" & RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
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
                RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.LCode In (10000,10001,10315) Order By TmpTrialBal.LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("DBal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                        "',' " & RsDataO.Fields("LName") & "'," & RsDataO.Fields("DBal") & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
                Set RsDataO = Nothing
                RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And TmpTrialBal.HCode=" & mGroup & " And TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.LCode Not In (10000,10001,10315) Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
                If LsvClient.TextMatrix(LsvClient.Row, 4) <> "Trust" Then
                    RsDataO.Open "Select IIF(ISNull(Sum(OpCr))=True,0,Sum(OpCr)) As RTotal From TmpTrialBal Where HCode In (58,13,5)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Else
                    RsDataO.Open "Select IIF(ISNull(Sum(OpCr))=True,0,Sum(OpCr)) As RTotal From TmpTrialBal Where HCode In (58,57,13,5)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                End If
            Else
                RsDataO.Open "Select IIF(ISNull(Sum(OpCr))=True,0,Sum(OpCr)) As RTotal From TmpTrialBal Where HCode In (58,57,5)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            If RsDataO.EOF = False Then
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & _
                    "','  Opening Balance'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
            End If
            Set RsDataO = Nothing
            If mFinYear = "19-20" Then
                If LsvClient.TextMatrix(LsvClient.Row, 4) <> "Trust" Then
                    RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName, Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where HCode In (5,13,58)" & _
                    "Group By IIF(Side=HSide,'Add','Less'),EName Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                Else
                    RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName, Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where HCode In (5,13,57,58)" & _
                    "And ECode Not In (960,961,962,964) Group By IIF(Side=HSide,'Add','Less'),EName Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                End If
            Else
                If LsvClient.TextMatrix(LsvClient.Row, 4) <> "Trust" Then
                    RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName,Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where HCode In (5,57,58)" & _
                    "Group By IIF(Side=HSide,'Add','Less'),EName Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                Else
                    RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName,Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where HCode In (5,57,58)" & _
                    "And ECode Not In (960,961,962,964) Group By IIF(Side=HSide,'Add','Less'),EName Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                End If
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
    .Formulas(13) = "mType='B'"
    .Formulas(14) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Action = 1
    For i = 0 To 14
        .Formulas(i) = ""
    Next
End With
mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
SetParent
End Sub
Private Sub BSheetPrint()
Dim i As Double
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpBSPrn"
With VsfHelp
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Insert Into TmpBSPrn (LName,LAmt,LNote,RName,RAmt,RNote,SrN) Values ('" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 1)) & _
        ",'" & .TextMatrix(i, 3) & "'," & Val(.TextMatrix(i, 5)) & "," & Val(.TextMatrix(i, 4)) & "," & i & ")"
    Next
End With
DbDataDB.CommitTrans
With RepPrint
    .Connect = MSCONNECT
    If LsvClient.TextMatrix(LsvClient.Row, 4) = "Trust" Then
        .ReportFileName = App.Path + "\Report\ABSRpt.Rpt"
        .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
        .Formulas(1) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
        .Formulas(2) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
        .Formulas(3) = "mTitle1='Balance Sheet as on " & RsRep.Fields("RtDt") & "'"
        .Formulas(4) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
        .Formulas(5) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
        .Formulas(6) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
        .Formulas(7) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
'        .Formulas(8) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
        .Formulas(9) = "mTRegNo='" & RsClient.Fields("TRegNo") & "'"
        .Formulas(10) = "mTRegDate='" & RsClient.Fields("TRegDt") & "'"
        .Formulas(11) = "mTAddress='" & RsClient.Fields("Address") & IIf(Len(RsClient.Fields("Address")) > 0, ", ", "") & RsClient.Fields("City") & ", " & ", " & RsClient.Fields("Taluka") & ", " & RsClient.Fields("District") & ", " & RsClient.Fields("State") & ", " & RsClient.Fields("PinCode") & "'"
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
    Else
        .ReportFileName = App.Path + "\Report\BSXRpt.Rpt"
        .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
        .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
        .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
        .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
        .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
        .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
        .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
        .Formulas(8) = "mSubHead=''"
        .Formulas(9) = "mTitle1='Balance Sheet as on " & RsRep.Fields("RtDt") & "'"
        .Formulas(10) = "mPlace='Place : " & RsRep.Fields("RPlace") & "'"
        .Formulas(11) = "mDate='Date : " & RsRep.Fields("RpDt") & "'"
        .Formulas(12) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
        .Formulas(13) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
        .Formulas(14) = "mTSign='Chief Functionary/Trustee'"
'        .Formulas(15) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
        .Formulas(16) = "mClient='" & LsvClient.TextMatrix(LsvClient.Row, 0) & "'"
        .Formulas(17) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
        .Formulas(18) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
        .Formulas(19) = "mLTitle='EXPENDITURE'"
        .Formulas(20) = "mRTitle='INCOMES'"
        .Action = 1
        For i = 0 To 20
            .Formulas(i) = ""
        Next
    End If
End With
End Sub
