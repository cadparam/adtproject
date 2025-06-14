VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmSchoolMemoTrn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "School Memo"
   ClientHeight    =   7164
   ClientLeft      =   540
   ClientTop       =   432
   ClientWidth     =   16428
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FSchMTrn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7164
   ScaleWidth      =   16428
   Begin VB.Frame FraNote 
      Height          =   6612
      Left            =   18000
      TabIndex        =   11
      Top             =   0
      Width           =   7692
      Begin VB.TextBox TxtOTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5640
         Width           =   1300
      End
      Begin VB.TextBox TxtNTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   5640
         Width           =   1300
      End
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   5640
         Width           =   1300
      End
      Begin VB.TextBox TxtNote 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   720
         TabIndex        =   2
         Top             =   5880
         Width           =   700
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   5400
         Picture         =   "FSchMTrn.frx":000C
         TabIndex        =   3
         ToolTipText     =   "Save"
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   6528
         Picture         =   "FSchMTrn.frx":0676
         TabIndex        =   4
         ToolTipText     =   "Cancel"
         Top             =   6120
         Width           =   975
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfNote 
         Height          =   5412
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7500
         _cx             =   13229
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
         FormatString    =   $"FSchMTrn.frx":0AB8
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
         ShowComboButton =   -1  'True
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
         Caption         =   "Note :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   5880
         Width           =   552
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   6492
         Index           =   2
         Left            =   0
         Top             =   120
         Width           =   7692
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
         FormatString    =   $"FSchMTrn.frx":0B01
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
         ShowComboButton =   -1  'True
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
      Height          =   7092
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   16212
      Begin VB.Frame FraDetail 
         Height          =   6132
         Left            =   18000
         TabIndex        =   16
         Top             =   0
         Width           =   7572
         Begin VB.TextBox TxtTotalD 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   5640
            Width           =   1536
         End
         Begin VB.CommandButton CmdCancelD 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   1248
            Picture         =   "FSchMTrn.frx":0B4A
            TabIndex        =   18
            ToolTipText     =   "Cancel"
            Top             =   5640
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfDetail 
            Height          =   5412
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   7380
            _cx             =   13017
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
            FormatString    =   $"FSchMTrn.frx":0F8C
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
            ShowComboButton =   -1  'True
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
            Width           =   7572
         End
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   4
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   732
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   7716
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   732
      End
      Begin VB.Frame FraMainExp 
         BackColor       =   &H00F7E0FE&
         Height          =   5412
         Left            =   18000
         TabIndex        =   26
         Top             =   840
         Width           =   15612
         Begin VB.CommandButton CmpMEClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Cancel"
            Top             =   4920
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfMainExport 
            Height          =   5052
            Left            =   0
            TabIndex        =   28
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
            FormatString    =   $"FSchMTrn.frx":0FD5
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
            ShowComboButton =   -1  'True
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
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export Data"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   1488
      End
      Begin VB.TextBox TxtCTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   6600
         Width           =   1776
      End
      Begin VB.TextBox TxtDTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   6600
         Width           =   1776
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   732
      End
      Begin VB.CommandButton CmdSearch 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1440
         Picture         =   "FSchMTrn.frx":101E
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
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp1 
         Height          =   5772
         Left            =   0
         TabIndex        =   22
         Top             =   720
         Width           =   7800
         _cx             =   13758
         _cy             =   10181
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
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FSchMTrn.frx":1654
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
         ShowComboButton =   -1  'True
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
         Height          =   5772
         Left            =   7920
         TabIndex        =   23
         Top             =   720
         Width           =   7800
         _cx             =   13758
         _cy             =   10181
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
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FSchMTrn.frx":169D
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
         ShowComboButton =   -1  'True
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
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   324
         Left            =   8040
         TabIndex        =   24
         Top             =   6600
         Width           =   60
      End
      Begin VB.Label LblCompany 
         Caption         =   "Client Name :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   10.8
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
         Height          =   6972
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   16212
      End
   End
   Begin Crystal.CrystalReport RepPrint 
      Left            =   0
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
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
Attribute VB_Name = "FrmSchoolMemoTrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mAcCode As Double
Dim mAuto As Double
Dim mLedCode As Double
Dim mActive As String
Private Sub CmdSearch_Click()
    FraClientHelp.Left = 180
    LsvClient.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then KeyAscii = 0
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub Form_Load()
    Me.Left = 50
    Me.Top = 50
    SetCombo
    SetData
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.CityName,'' As ParentNm,GroupMst.GroupName,AcMst.AcCode,'' As RName From AcMst,GroupMst Where AcMst.AcType=2 And " & _
"IsNull(AcMst.ParentCd)=True And AcMst.AcType=GroupMst.GroupCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.CityName,AcMst1.AcName+IIF(Len(AcMst1.CityName)>0,', '+AcMst1.CityName,'')," & _
"GroupMst.GroupName,AcMst.AcCode,'' From AcMst,GroupMst,AcMst As AcMst1 Where  AcMst.AcType=2 And AcMst.ParentCd=AcMst1.AcCode And AcMst.AcType=GroupMst.GroupCode Order By AcMst.FileNo", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
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
    RsQry.Open "Select AcMst.AcCode,AcMst2.AcName+', '+AcMst2.CityName As RName From AcMst,AcMst As AcMst1,AcMst As AcMst2 Where " & _
    "AcMst.AcType=2 And AcMst.ParentCd=AcMst1.AcCode And AcMst1.ParentCd=AcMst2.AcCode Order By AcMst.FileNo", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .Row = 1
        .Row = .FindRow(RsQry.Fields("AcCode"), , 5)
        If .Row > 1 Then .TextMatrix(.Row, 6) = RsQry.Fields("RName")
        RsQry.MoveNext
    Loop
End With
End Function
Private Sub ClearText()
Dim ObjText As Object
For Each ObjText In Me
    If TypeOf ObjText Is TextBox Then ObjText.Text = ""
Next
mAcCode = 0
mAuto = 0
mLedCode = 0
mActive = ""
FraClientHelp.Left = 18000
mActive = ""
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
    SetData
    ShowData
    SetTotal
    VsfHelp.SetFocus
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Save"
        If TxtName.Text = "" Then
            MsgBox "Sorry !! Not Allowded..", vbInformation, "Black Data Error"
            TxtName.SetFocus
        Else
            If MsgBox("Are you sure to save ? ", vbInformation + vbYesNo, "Confirmation") = vbYes Then
                SaveData
                ClearText
                SetCombo
                FraClientHelp.Left = 18000
                CmdSearch.SetFocus
            End If
        End If
    Case "Cancel"
        If MsgBox("Are you sure to Cancel ? ", vbInformation + vbYesNo, "Confirmation") = vbYes Then
            SetData
            ClearText
            SetCombo
            FraClientHelp.Left = 18000
            CmdSearch.SetFocus
        End If
    Case "Print"
        PrintRec
    Case "Exit"
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(1).Enabled = True
    TlbSav(2).Enabled = True
End Function
Private Sub SetData()
Dim RsQry As New ADODB.Recordset
With VsfHelp
    .Cols = 4
    .Rows = 1
    .Rows = 2
    .TextMatrix(0, 0) = "PAYMENTS"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "NOTE"
    .ColWidth(1) = 700
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1800
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .ColWidth(3) = 0    'CSIDE
    .Col = 1
    .Row = 1
    .Refresh
End With
With VsfHelp1
    .Cols = 4
    .Rows = 1
    .Rows = 2
    .TextMatrix(0, 0) = "RECEIPTS"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "NOTE"
    .ColWidth(1) = 700
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1800
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .ColWidth(3) = 0    'DSIDE
    .Col = 1
    .Row = 1
    .Refresh
End With
End Sub

Private Sub VsfHelp_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SetTotal
End Sub
Private Sub SetTotal()
Dim i As Double
TxtCTotal.Text = "0.00"
TxtDTotal.Text = "0.00"
With VsfHelp1
    .TextMatrix(13, 2) = "0.00"
    .TextMatrix(21, 2) = "0.00"
    For i = 7 To .Rows - 1
        .TextMatrix(13, 2) = Val(.TextMatrix(13, 2)) + Val(.TextMatrix(i, 2))
        If i = 12 Then Exit For
    Next
    For i = 16 To .Rows - 1
        .TextMatrix(21, 2) = Val(.TextMatrix(21, 2)) + Val(.TextMatrix(i, 2))
        If i = 20 Then Exit For
    Next
    TxtDTotal.Text = Val(.TextMatrix(4, 2)) + Val(.TextMatrix(13, 2)) + Val(.TextMatrix(21, 2))
End With
With VsfHelp
    .TextMatrix(10, 2) = "0.00"
    .TextMatrix(17, 2) = "0.00"
    For i = 2 To .Rows - 1
        .TextMatrix(10, 2) = Val(.TextMatrix(10, 2)) + Val(.TextMatrix(i, 2))
        If i = 9 Then Exit For
    Next
    For i = 13 To .Rows - 1
        .TextMatrix(17, 2) = Val(.TextMatrix(17, 2)) + Val(.TextMatrix(i, 2))
        If i = 16 Then Exit For
    Next
    TxtCTotal.Text = Val(.TextMatrix(10, 2)) + Val(.TextMatrix(17, 2)) + Val(.TextMatrix(22, 2))
End With
TxtCTotal.Text = Format(TxtCTotal.Text, "0.00")
TxtDTotal.Text = Format(TxtDTotal.Text, "0.00")
If Val(TxtCTotal.Text) - Val(TxtDTotal.Text) = 0 Then LblTotal.Caption = "" Else LblTotal.Caption = "Dif.Rs.-->" & Format(CStr(Abs(Val(TxtCTotal.Text) - Val(TxtDTotal.Text))), "0.00")
End Sub
Private Sub ShowData()
Dim RsQ As New ADODB.Recordset
Dim i As Double
With VsfHelp
    .Col = 2
    .Row = 1
    .TextMatrix(.Row, 0) = "Direct Recurring Expenditure"
    .Cell(flexcpFontBold, .Row, 0, .Row, 0) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbBlue
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & _
    mAcCode & " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=81", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Salary - Staff of Teachers"
    If RsQ.Fields("RTotal") <> 0 Then .TextMatrix(.Row, 2) = RsQ.Fields("RTotal") Else .TextMatrix(.Row, 2) = "0.00"
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & _
    mAcCode & " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=67", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Salary - Non-Teaching Staff"
    If RsQ.Fields("RTotal") <> 0 Then .TextMatrix(.Row, 2) = RsQ.Fields("RTotal") Else .TextMatrix(.Row, 2) = "0.00"
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amount))=True,0,Sum(Amount)) As RTotal From PlDtl Where SrN<9999 And AcCode=" & mAcCode & " And SLedCode=121", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Rent, Taxes and Insurance"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & _
    mAcCode & " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=38", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 2) = Val(.TextMatrix(.Row, 2)) + RsQ.Fields("RTotal")
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & _
    mAcCode & " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=230", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Office Contingencies"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amount))=True,0,Sum(Amount)) As RTotal From PlDtl Where SrN<9999 And AcCode=" & mAcCode & " And SLedCode In (130,132)", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Books & Prizes"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst,LedgerMst Where PlDtl.AcCode=" & _
    mAcCode & " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=LedgerMst.LedCode And LedgerMst.LedCode<>72 And LedgerMst.GroupCode=7", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Furniture & Equipment for which a special grant has not been claimed"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    .Rows = .Rows + 1
    .Row = .Rows - 1
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=39", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 0) = "Current Repairs"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    .Rows = .Rows + 1
    .Row = .Rows - 1
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst,LedgerMst,HeadMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.SLedCode Not In (121,130,132) And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=LedgerMst.LedCode And " & _
    "LedgerMst.LedCode Not In (67,81,84,230) And LedgerMst.GroupCode=HeadMst.GroupCode And HeadMst.GType='I' And HeadMst.Active='D' And HeadMst.GroupCode Not In (7,10,12,14,19,20,22)", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 0) = "Miscellanous Expenses(65%)"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Total Recurring Expenditure"
    .Cell(flexcpFontBold, .Row, 0, .Row, 2) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 2) = vbBlue
    .Rows = .Rows + 2
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Indirect & Non-Recurring Expenditure"
    .Cell(flexcpFontBold, .Row, 0, .Row, 0) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbBlue
    .Rows = .Rows + 1
    .Row = .Rows - 1
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=84", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 0) = "Scholarships/Other Payments"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(PlDtl.TrnType=HeadMst.Active,PlDtl.Amount,PlDtl.Amount*-1)) As RTotal From PlDtl,SubLedMst,LedgerMst,HeadMst Where PlDtl.AcCode=" & _
    mAcCode & " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=LedgerMst.LedCode And LedgerMst.GroupCode In (10,14)" & _
    " And HeadMst.GroupCode=LedgerMst.GroupCode", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Transfer to Reserve Fund/Investments"
    .TextMatrix(.Row, 2) = IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal"))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=72", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Furniture & Equipment for which a special grant has been claimed"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amount))=True,0,Sum(Amount)) As RTotal From PlDtl Where AcCode=" & mAcCode & " And SrN<9999 And SLedCode=13", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Repayment of Management Loans"
    .TextMatrix(.Row, 2) = RsQ.Fields("RTotal")
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Total Non-Recurring Expenditure"
    .Cell(flexcpFontBold, .Row, 0, .Row, 2) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 2) = vbBlue
    .Rows = .Rows + 2
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Balance as on 31st March 20" & Mid(RsComp.Fields("FinYr"), 4, 2)
    .Cell(flexcpFontBold, .Row, 0, .Row, 0) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbBlue
    Set RsQ = Nothing
    RsQ.Open "Select * From PlDtl Where AcCode=" & mAcCode & " And SrN>=9999 And SLedCode=4", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Cash"
    If RsQ.EOF = False Then .TextMatrix(.Row, 2) = RsQ.Fields("Amount") Else .TextMatrix(.Row, 2) = "0.00"
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amount))=True,0,Sum(Amount)) As RTotal From PlDtl Where AcCode=" & mAcCode & " And SrN>=9999 And SLedCode<>4", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Bank"
    If RsQ.Fields("RTotal") <> 0 Then .TextMatrix(.Row, 2) = RsQ.Fields("RTotal") Else .TextMatrix(.Row, 2) = "0.00"
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Total Cash / Bank"
    .Cell(flexcpFontBold, .Row, 0, .Row, 2) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 2) = vbBlue
    If (.TextMatrix(.Row - 1, 0) = "Cash" Or .TextMatrix(.Row - 1, 0) = "Bank") Then .TextMatrix(.Row, 2) = Val((.TextMatrix(.Row - 1, 2)))
    If .TextMatrix(.Row - 2, 0) = "Cash" Then .TextMatrix(.Row, 2) = Val(.TextMatrix(.Row, 2)) + Val((.TextMatrix(.Row - 2, 2)))
    .Refresh
    For i = 1 To .Rows - 1
        .TextMatrix(i, 3) = i
    Next
    Set RsQ = Nothing
    RsQ.Open "Select * From SchoolMemoMst Where AcCode=" & mAcCode & " and AcSide='D' Order By HeadCode", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        i = .FindRow(RsQ.Fields("HeadCode"), 1, 3)
        If i > 0 Then .TextMatrix(i, 1) = RsQ.Fields("AcNote")
        RsQ.MoveNext
    Loop
End With
With VsfHelp1
    Set RsQ = Nothing
    RsQ.Open "Select AcBal.AcOpen From AcBal Where AcCode=" & mAcCode & " And LedCode=4", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 0) = "Balance as on 1st April 20" & Mid(RsComp.Fields("FinYr"), 1, 2)
    .Cell(flexcpFontBold, .Row, 0, .Row, 0) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbBlue
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Cash"
    .TextMatrix(.Row, 2) = IIf(RsQ.EOF = False, RsQ.Fields("AcOpen"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(AcBal.AcOpen))=True,0,Sum(AcBal.AcOpen)) As RTotal From AcBal,LedgerMst Where AcBal.AcCode=" & _
    mAcCode & " And AcBal.LedCode<>4 And AcBal.LedCode=LedgerMst.LedCode And LedgerMst.GroupCode=2", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Bank"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Cash / Bank Total"
    .TextMatrix(.Row, 2) = Val(.TextMatrix(2, 2)) + Val(.TextMatrix(3, 2))
    .Cell(flexcpFontBold, .Row, 0, .Row, 2) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 2) = vbBlue
    .Rows = .Rows + 2
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Recurring Receipts"
    .Cell(flexcpFontBold, .Row, 0, .Row, 0) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbBlue
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=80", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Provincial Grants"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=29", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Fees and Fines"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=227", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Subscriptions"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst,LedgerMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode Not In (29,227,228,229,283) And SubLedMst.LedCode=" & _
    "LedgerMst.LedCode And LedgerMst.GroupCode In (43,44,45,46,47,48,52)", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Income from Other Sources"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=228", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Endowment for Maintenance of School"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=229", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Nominal Receipts"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Total Recurring Receipts"
    .Cell(flexcpFontBold, .Row, 0, .Row, 2) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 2) = vbBlue
    .Rows = .Rows + 2
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Non-Recurring Receipts"
    .Cell(flexcpFontBold, .Row, 0, .Row, 0) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbBlue
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=96", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Provincial Grants"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amount))=True,0,Sum(Amount)) As RTotal From PlDtl Where AcCode=" & mAcCode & _
    " And SrN<9999 And TrnType='C' And SLedCode=12", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Loans from Management"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst,LedgerMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=LedgerMst.LedCode And LedgerMst.GroupCode=1", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Refund of PF/Advances"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode in (16,283)", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Donations/Collection for Specific Purpose"
    .TextMatrix(.Row, 2) = IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst,LedgerMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=LedgerMst.LedCode And LedgerMst.GroupCode=49", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(PlDtl.Amount))=True,0,Sum(PlDtl.Amount)) As RTotal From PlDtl,SubLedMst,LedgerMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=LedgerMst.LedCode And LedgerMst.GroupCode=49", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
     .TextMatrix(.Row, 2) = Val(.TextMatrix(.Row, 2)) + IIf(RsQ.Fields("RTotal") <> 0, RsQ.Fields("RTotal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(PlDtl.TrnType='C',PlDtl.Amount,PlDtl.Amount*-1)) As RTotal From PlDtl,SubLedMst,LedgerMst Where PlDtl.AcCode=" & mAcCode & _
    " And PlDtl.SrN<9999 And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=LedgerMst.LedCode And LedgerMst.GroupCode=11", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Liabilities"
    .TextMatrix(.Row, 2) = IIf(IsNull(RsQ.Fields("RTotal")) = False, RsQ.Fields("RTotal"), 0)
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Total Non-Recurring Receipts"
    .Cell(flexcpFontBold, .Row, 0, .Row, 2) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 2) = vbBlue
    .Refresh
    For i = 1 To .Rows - 1
        .TextMatrix(i, 3) = i
    Next
    Set RsQ = Nothing
    RsQ.Open "Select * From SchoolMemoMst Where AcCode=" & mAcCode & " and AcSide='C' Order By HeadCode", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        i = .FindRow(RsQ.Fields("HeadCode"), 1, 3)
        If i > 0 Then .TextMatrix(i, 1) = RsQ.Fields("AcNote")
        RsQ.MoveNext
    Loop
End With
End Sub
Private Sub ExpMainData()
Dim i As Double
With VsfMainExport
    .Cols = 6
    .Rows = 1
    .TextMatrix(0, 0) = "RECEIPTS"
    .ColWidth(0) = 3000
    .TextMatrix(0, 1) = "NOTE"
    .ColWidth(1) = 1000
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "PAYMENTS"
    .ColWidth(3) = 3000
    .TextMatrix(0, 4) = "NOTE"
    .ColWidth(4) = 3000
    .TextMatrix(0, 5) = "AMOUNT RS"
    .ColWidth(5) = 1500
    .ColFormat(5) = "0.00"
    .ColAlignment(5) = flexAlignRightCenter
    .Refresh
    .Rows = .Rows + 1
    .Row = 1
    For i = 1 To VsfHelp1.Rows - 1
        .TextMatrix(i, 0) = VsfHelp1.TextMatrix(i, 0)
        .TextMatrix(i, 1) = VsfHelp1.TextMatrix(i, 1)
        .TextMatrix(i, 2) = VsfHelp1.TextMatrix(i, 2)
        .Rows = .Rows + 1
    Next
    For i = 1 To VsfHelp.Rows - 1
        .TextMatrix(i, 3) = VsfHelp.TextMatrix(i, 0)
        .TextMatrix(i, 4) = VsfHelp.TextMatrix(i, 1)
        .TextMatrix(i, 5) = VsfHelp.TextMatrix(i, 2)
        If i = .Rows - 1 Then .Rows = .Rows + 1
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
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 2) = TxtDTotal.Text
    .TextMatrix(.Row, 5) = TxtCTotal.Text
End With
If Dir(App.Path + "\Excel", vbDirectory) = "" Then MkDir App.Path + "\Excel"
VsfMainExport.SaveGrid App.Path & "\EXCEL\SCHOOLMEMO.XLS", flexFileTabText
MsgBox "Successfully Excel File Generated In " + App.Path + "\EXCEL\SCHOOLMEMO.XLS", vbInformation, "Alert"
End Sub
Private Sub CmdExport_Click()
    'If mUType = "Admin" Then ExpMainData
    ExpMainData
End Sub
Private Function SaveData()
Dim i As Double
On Error GoTo XErr
DbWorkAuto.BeginTrans
DbWorkAuto.Execute "Delete From SchoolMemoMst Where AcCode=" & mAcCode
With VsfHelp
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 Then
            DbWorkAuto.Execute "Insert InTo SchoolMemoMst (AcCode,HeadCode,AcNote,AcSide) Values (" & mAcCode & "," & i & "," & Val(.TextMatrix(i, 1)) & ",'D')"
        End If
    Next
End With
With VsfHelp1
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 Then
            DbWorkAuto.Execute "Insert InTo SchoolMemoMst (AcCode,HeadCode,AcNote,AcSide) Values (" & mAcCode & "," & i & "," & Val(.TextMatrix(i, 1)) & ",'C')"
        End If
    Next
End With
DbWorkAuto.CommitTrans
MsgBox "Record Succussfully Saved.", vbInformation, "Alert"
Exit Function
XErr:
MsgBox Err.Description
DbWorkAuto.RollbackTrans
End Function
Private Sub VsfHelp_EnterCell()
With VsfHelp
    If .Col = 1 Then
        .Editable = flexEDKbd
        On Error Resume Next
        SendKeys "{F2}"
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub VsfHelp_RowColChange()
With VsfHelp
    If .Col = 1 Then
        .Editable = flexEDKbd
        On Error Resume Next
        SendKeys "{F2}"
    Else
        .Editable = flexEDNone
    End If
End With
End Sub

Private Sub VsfHelp1_DblClick()
    SetLedger VsfHelp1.Row, "C"
    If VsfDetail.TextMatrix(1, 0) <> "" Then
        FraDetail.Left = IIf(Me.ActiveControl.Name = "VsfHelp1", 8000, 180)
        VsfDetail.SetFocus
    End If
End Sub

Private Sub VsfHelp1_EnterCell()
With VsfHelp1
    If .Col = 1 Then
        .Editable = flexEDKbd
        On Error Resume Next
        SendKeys "{F2}"
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub VsfHelp1_RowColChange()
With VsfHelp1
    If .Col = 1 Then
        .Editable = flexEDKbd
        On Error Resume Next
        SendKeys "{F2}"
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub SetLedger(ByVal mRow As Integer, ByVal mSide As String)
Dim RsQry As New ADODB.Recordset
With VsfDetail
    .Editable = flexEDNone
    .Cols = 3
    .Rows = 1
    .Row = 0
    .TextMatrix(0, 0) = "LEDGER"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "AMOUNT"
    .ColWidth(1) = 1700
    .ColFormat(1) = "0.00"
    .ColAlignment(1) = flexAlignRightCenter
    .FixedAlignment(1) = flexAlignRightCenter
    .ColWidth(2) = 100  '  LEDGERCODE
    .Col = 0
    .Refresh
    .Rows = 2
    .Row = 1
    If mSide = "C" Then
        Select Case mRow
            Case 3  '   Bank Opening
                Set RsQry = Nothing
                RsQry.Open "Select LedgerMst.LedCode,LedgerMst.LedName,AcBal.AcOpen As RTotal From AcBal,LedgerMst Where AcBal.AcCode=" & _
                mAcCode & " And AcBal.LedCode<>4 And AcBal.LedCode=LedgerMst.LedCode And LedgerMst.GroupCode=2 Order By LedgerMst.LedName", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
                TxtTotalD.Text = "0.00"
                Do While RsQry.EOF = False
                    .TextMatrix(.Row, 0) = RsQry.Fields("LedName")
                    .TextMatrix(.Row, 1) = RsQry.Fields("RTotal")
                    .TextMatrix(.Row, 2) = RsQry.Fields("LedCode")
                    TxtTotalD.Text = Val(TxtTotalD.Text) + RsQry.Fields("RTotal")
                    RsQry.MoveNext
                    If RsQry.EOF = False Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                Loop
            Case 7  '   Provincial Grants under Recurring Receipts
                Set RsQry = Nothing
                RsQry.Open "Select SubLedMst.SLedCode,SubLedMst.SLedName,PlDtl.Amount As RTotal From PlDtl,SubLedMst Where PlDtl.AcCode=" & _
                mAcCode & "And PlDtl.SrN<9999 And PlDtl.TrnType='C' And PlDtl.SLedCode=SubLedMst.SLedCode And SubLedMst.LedCode=80", DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
                TxtTotalD.Text = "0.00"
                Do While RsQry.EOF = False
                    .TextMatrix(.Row, 0) = RsQry.Fields("SLedName")
                    .TextMatrix(.Row, 1) = RsQry.Fields("RTotal")
                    .TextMatrix(.Row, 2) = RsQry.Fields("SLedCode")
                    TxtTotalD.Text = Val(TxtTotalD.Text) + RsQry.Fields("RTotal")
                    RsQry.MoveNext
                    If RsQry.EOF = False Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                Loop
            Case 8  '   Fees and Fines
            Case 9  '   Subscriptions
            Case 10 '   Income from Other Sources
            Case 11 '   Endowment for Maintenance of School
            Case 12 '   Nominal Receipts
            Case 16 '   Provincial Grants Under Non-Recurring Receipts
            Case 17 '   Loans from Management
            Case 18 '   Refund of PF/Advances
            Case 19 '   Donations for Specific Purpose
            Case 20 '   Liabilities
        End Select
    End If
End With
TxtTotalD.Text = Format(TxtTotalD.Text, "0.00")
End Sub
Private Sub CmdCancelD_Click()
    FraDetail.Left = 18000
End Sub
Private Sub PrintRec()
Dim i As Double
Dim RsRep As New ADODB.Recordset
RsRep.Open "Select * From AnlRepMst Where AcCode=" & mAcCode, DbWorkAuto, adOpenDynamic, adLockReadOnly, adCmdText
If RsRep.EOF = False Then
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        MsgBox "UDIN not issued.", vbCritical, "Alert"
        Exit Sub
    End If
End If
DbWorkAuto.BeginTrans
DbWorkAuto.Execute "Delete From TmpSMPrn"
        'DbWorkAuto.Execute "Insert Into TmpSMPrn (LName,LAmt,LNote,RName,RAmt,RNote,SrN) Values ('" & .TextMatrix(i, 0) & "'," & _
        Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 3) & "'," & Val(.TextMatrix(i, 5)) & "," & _
        Val(.TextMatrix(i, 4)) & "," & i & ")"

With VsfHelp1
    For i = 1 To .Rows - 1
        DbWorkAuto.Execute "Insert Into TmpSMPrn (LName,LAmt,LNote,SrN) Values ('" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & _
        "," & Val(.TextMatrix(i, 1)) & "," & i & ")"
    Next
End With
With VsfHelp
    For i = 1 To .Rows - 1
        If i < VsfHelp1.Rows Then
            DbWorkAuto.Execute "Update TmpSMPrn Set RName='" & .TextMatrix(i, 0) & "',RAmt=" & Val(.TextMatrix(i, 2)) & ",RNote=" & _
            Val(.TextMatrix(i, 1)) & " Where SrN=" & i
        Else
            DbWorkAuto.Execute "Insert Into TmpSMPrn (RName,RAmt,RNote,SrN) Values ('" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & _
            "," & Val(.TextMatrix(i, 1)) & "," & i & ")"
        End If
    Next
End With
DbWorkAuto.CommitTrans
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\5SMemo.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("CompName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("TinNo") & "'"
    .Formulas(6) = "mTitle='(FRN: " & RsComp.Fields("StateName") & ")'"
    If Len(LsvClient.TextMatrix(LsvClient.Row, 6)) > 0 Then
        .Formulas(7) = "mHead='" & Mid(LsvClient.TextMatrix(LsvClient.Row, 6), 1, Len(LsvClient.TextMatrix(LsvClient.Row, 6)) - 9) & "'"
    ElseIf Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0 Then
        .Formulas(7) = "mHead='" & Mid(LsvClient.TextMatrix(LsvClient.Row, 3), 1, Len(LsvClient.TextMatrix(LsvClient.Row, 3)) - 9) & "'"
    Else
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + Mid(LsvClient.TextMatrix(LsvClient.Row, 2), 1, Len(LsvClient.TextMatrix(LsvClient.Row, 2)) - 9) & "'"
    End If
    If RsRep.EOF = False Then
        .Formulas(8) = "mUDIN='UDIN: " & RsRep.Fields("RUDIN") & "'"
        .Formulas(9) = "mTitle1='School Memo as on " & RsRep.Fields("RTDate") & "'"
        .Formulas(10) = "mPlace='Place : " & RsRep.Fields("SPlace") & "'"
        .Formulas(11) = "mDate='Date : " & RsRep.Fields("RDate") & "'"
        .Formulas(12) = "mAuditNm='" & RsRep.Fields("SAuditName") & "'"
        .Formulas(13) = "mAuditNo='Mem.No. : " & RsRep.Fields("SAudit") & "'"
    Else
        .Formulas(8) = "mUDIN=''"
        .Formulas(9) = "mTitle1=''"
        .Formulas(10) = "mPlace='Place :'"
        .Formulas(11) = "mDate='Date :'"
        .Formulas(12) = "mAuditNm=''"
        .Formulas(13) = "mAuditNo='Mem.No. :'"
    End If
    .Formulas(14) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(15) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Action = 1
    For i = 0 To 15
        .Formulas(i) = ""
    Next
End With
End Sub
