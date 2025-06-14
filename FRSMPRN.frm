VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmSchoolMemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "School Memo"
   ClientHeight    =   8880
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   16425
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRSMPRN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   16425
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
      Index           =   4
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Print"
      Top             =   240
      Width           =   735
   End
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
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
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
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   5400
         Picture         =   "FRSMPRN.frx":000C
         TabIndex        =   3
         ToolTipText     =   "Save"
         Top             =   6120
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
         Height          =   372
         Index           =   1
         Left            =   6528
         Picture         =   "FRSMPRN.frx":0676
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
         FormatString    =   $"FRSMPRN.frx":0AB8
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
         Caption         =   "Note :"
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
         FormatString    =   $"FRSMPRN.frx":0B01
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
      Height          =   8775
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
            TabIndex        =   19
            Top             =   5640
            Width           =   1536
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
            Picture         =   "FRSMPRN.frx":0B4A
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
            FormatString    =   $"FRSMPRN.frx":0F8C
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
            Width           =   7572
         End
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
         Index           =   3
         Left            =   6876
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   732
      End
      Begin VB.Frame FraMainExp 
         BackColor       =   &H00F7E0FE&
         Height          =   5412
         Left            =   18000
         TabIndex        =   25
         Top             =   840
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
            FormatString    =   $"FRSMPRN.frx":0FD5
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
         Left            =   13560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   8280
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   8280
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   732
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
         Picture         =   "FRSMPRN.frx":101E
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
         Height          =   7500
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   7900
         _cx             =   13935
         _cy             =   13229
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
         FormatString    =   $"FRSMPRN.frx":1654
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
         Height          =   7500
         Left            =   8160
         TabIndex        =   23
         Top             =   720
         Width           =   7905
         _cx             =   13935
         _cy             =   13229
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
         FormatString    =   $"FRSMPRN.frx":169D
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
         Left            =   8280
         TabIndex        =   24
         Top             =   8280
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
         Height          =   8655
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   16215
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
Attribute VB_Name = "FrmSchoolMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim mAcCode As Double
Dim mAuto As Double
Dim mLedCode As Double
Dim mActive As String
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
    SetData
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode,'' As RName From AcMst,GrpMst Where AcMst.AcType=3 And " & _
"AcMst.Active=-1 And AcMst.PaCode=0 And AcMst.AcType=GrpMst.GCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.City,AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,'')," & _
"GrpMst.GName,AcMst.AcCode,'' From AcMst,GrpMst,AcMst As AcMst1 Where  AcMst.AcType=3 And AcMst.PaCode=AcMst1.AcCode And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    RsQry.Open "Select AcMst.AcCode,AcMst2.AcName+', '+AcMst2.City As RName From AcMst,AcMst As AcMst1,AcMst As AcMst2 Where AcMst.AcType=3 And " & _
    "AcMst.PaCode=AcMst1.AcCode And AcMst1.PaCode=AcMst2.AcCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    Dim RsQ As New ADODB.Recordset
    DBWorkTmp.BeginTrans
    DBWorkTmp.Execute "Delete From TmpTrialBal"
    DBWorkTmp.Execute "Delete From TmpCtDtl"
    DBWorkTmp.Execute "Delete From TmpRpDtl"
    DBWorkTmp.CommitTrans
    DBWorkTmp.BeginTrans
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
        DBWorkTmp.Execute "Insert InTo TmpCtDtl (HType,HSide,AcCode,HCode,LCode,ECode,EName,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
        RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("EName") & "','" & RsQ.Fields("Side") & _
        "'," & RsQ.Fields("Amt") & ")"
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
    SetData
    ShowData
    SetTotal
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHOOL_MEMO','VIEW','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    VsfHelp.SetFocus
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
                SaveData
                DbDataDB.BeginTrans
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHOOL_MEMO','SAVE_DATA','" & Date & "','" & Time & "')"
                DbDataDB.CommitTrans
                ClearText
                SetCombo
                FraClientHelp.Left = 18000
                CmdSearch.SetFocus
            End If
        End If
    Case "Print"
        If MsgBox("Save data and Print?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
            SaveData
            DbDataDB.BeginTrans
'                DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHOOL_MEMO','SAVE_DATA','" & Date & "','" & Time & "')"
            DbDataDB.CommitTrans
        End If
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHOOL_MEMO','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
        ClearText
    Case "Cancel"
        If MsgBox("Are you sure to Cancel? ", vbInformation + vbYesNo, "Confirmation") = vbYes Then
            SetData
            ClearText
            SetCombo
            FraClientHelp.Left = 18000
            CmdSearch.SetFocus
        End If
    Case "Exit"
        If MsgBox("Close All Reports? ", vbInformation + vbYesNo, "Confirmation") = vbYes Then
            Set DBWorkTmp = Nothing
            Unload Me
        End If
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
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=61", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Salary - Staff of Teachers"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=60", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Salary - Non-Teaching Staff"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where (ECode In (485,999) or LCode=43 or HCode=19)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Rent, Taxes and Insurance"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=63", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Office Contingencies"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where ECode In (549,484)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Books & Prizes"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode<>14 And HCode=8", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Furniture & Equipment for which a special grant has not been claimed"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    .Rows = .Rows + 1
    .Row = .Rows - 1
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where HCode=17", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 0) = "Current Repairs"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    .Rows = .Rows + 1
    .Row = .Rows - 1
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where HSide='D' And ECode Not In (484,485,549,999) And LCode Not In (5,14,43,60,61,62,63) And HCode Not In (2,7,8,10,11,14,17,19,24,26,28,29,30,31,32,33)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 0) = "Miscellanous Expenses(65%)"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
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
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where (HCode In (24,26,28,29,30,31,32,33) Or LCode=62) And ECode Not In (485,999)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 0) = "Scholarships/Other Payments"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where HCode In (7,2)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Transfer to Reserve Fund/Investments"
    .TextMatrix(.Row, 2) = IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal"))
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=14", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Furniture & Equipment for which a special grant has been claimed"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='D',Amt,0)) As RTotal From TmpRpDtl Where HCode=14", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Repayment of Management Loans"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
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
    RsQ.Open "Select * From RpDtl Where AcCode=" & mAcCode & " And SrN>=9999 And ECode=10000", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Cash"
    If RsQ.EOF = False Then .TextMatrix(.Row, 2) = RsQ.Fields("Amt") Else .TextMatrix(.Row, 2) = "0.00"
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From RpDtl Where AcCode=" & mAcCode & " And SrN>=9999 And ECode<>10000", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Bank"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Cash/Bank Total"
    .Cell(flexcpFontBold, .Row, 0, .Row, 2) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 2) = vbBlue
    If (.TextMatrix(.Row - 1, 0) = "Cash" Or .TextMatrix(.Row - 1, 0) = "Bank") Then .TextMatrix(.Row, 2) = Val((.TextMatrix(.Row - 1, 2)))
    If .TextMatrix(.Row - 2, 0) = "Cash" Then .TextMatrix(.Row, 2) = Val(.TextMatrix(.Row, 2)) + Val((.TextMatrix(.Row - 2, 2)))
    .Refresh
    For i = 1 To .Rows - 1
        .TextMatrix(i, 3) = i
    Next
    Set RsQ = Nothing
    RsQ.Open "Select * From NtDtl Where AcCode=" & mAcCode & " And RType=3 And RSide='D' Order By HCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        i = .FindRow(RsQ.Fields("HCode"), 1, 3)
        If i > 0 Then .TextMatrix(i, 1) = RsQ.Fields("Note")
        RsQ.MoveNext
    Loop
End With
With VsfHelp1
    Set RsQ = Nothing
    RsQ.Open "Select OpDtl.OpBal From OpDtl Where AcCode=" & mAcCode & " And LCode=10000", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(.Row, 0) = "Balance as on 1st April 20" & Mid(RsComp.Fields("FinYr"), 1, 2)
    .Cell(flexcpFontBold, .Row, 0, .Row, 0) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbBlue
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Cash"
    .TextMatrix(.Row, 2) = IIf(RsQ.EOF = False, RsQ.Fields("OpBal"), 0)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(OpDtl.OpBal))=True,0,Sum(OpDtl.OpBal)) As RTotal From OpDtl,LedMst Where OpDtl.AcCode=" & mAcCode & _
    " And OpDtl.LCode<>10000 And OpDtl.LCode=LedMst.LCode And LedMst.HCode=12", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Bank"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Cash/Bank Total"
    .TextMatrix(.Row, 2) = Val(.TextMatrix(2, 2)) + Val(.TextMatrix(3, 2))
    .Cell(flexcpFontBold, .Row, 0, .Row, 2) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 2) = vbBlue
    .Rows = .Rows + 2
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Recurring Receipts"
    .Cell(flexcpFontBold, .Row, 0, .Row, 0) = True
    .Cell(flexcpForeColor, .Row, 0, .Row, 0) = vbBlue
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=73", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Provincial Grants"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=78", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Fees and Fines"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=79", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Subscriptions"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where HCode In (42,44,45,46,47,48,53) And LCode Not In (78,79,80,81,82)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Income from Other Sources"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=80", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Endowment for Maintenance of School"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode=81", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Nominal Receipts"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
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
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where LCode In (74,83)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Provincial Grants"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,0)) As RTotal From TmpRpDtl Where HCode=14", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Loans from Management"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where HCode In (10,11)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Refund of PF/Advances"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where (HCode=50 or (LCode=82 or LCode=72))", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 0) = "Donations/Collection for Specific Purpose"
    .TextMatrix(.Row, 2) = Format(IIf(IsNull(RsQ.Fields("RTotal")) = True, 0, RsQ.Fields("RTotal")), "0.00")
    Set RsQ = Nothing
    RsQ.Open "Select Sum(IIF(Side='C',Amt,Amt*-1)) As RTotal From TmpRpDtl Where HCode In (3,4)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    RsQ.Open "Select * From NtDtl Where AcCode=" & mAcCode & " And RType=3 And RSide='C' Order By HCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        i = .FindRow(RsQ.Fields("HCode"), 1, 3)
        If i > 0 Then .TextMatrix(i, 1) = RsQ.Fields("Note")
        RsQ.MoveNext
    Loop
End With
End Sub
Private Sub CmdCancelD_Click()
    FraDetail.Left = 18000
End Sub
Private Function SaveData()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From NtDtl Where RType=3 And AcCode=" & mAcCode
With VsfHelp
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 Then
            DbDataDB.Execute "Insert InTo NtDtl (AcCode,HCode,[Note],RSide,RType) Values (" & mAcCode & "," & i & "," & Val(.TextMatrix(i, 1)) & ",'D',3)"
        End If
    Next
End With
With VsfHelp1
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 Then
            DbDataDB.Execute "Insert InTo NtDtl (AcCode,HCode,[Note],RSide,RType) Values (" & mAcCode & "," & i & "," & Val(.TextMatrix(i, 1)) & ",'C',3)"
        End If
    Next
End With
DbDataDB.CommitTrans
MsgBox "Record Successfully Saved.", vbInformation, "Alert"
Exit Function
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
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
Private Sub VsfHelp_DblClick()
    SetLedger VsfHelp.Row, "D"
    If VsfDetail.TextMatrix(1, 0) <> "" Then
        FraDetail.Left = IIf(Me.ActiveControl.Name = "VsfHelp", 8000, 180)
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
                RsQry.Open "Select LedMst.LCode,LedMst.LName,OpDtl.OpBal As RTotal From OpDtl,LedMst Where OpDtl.AcCode=" & mAcCode & " And OpDtl.LCode<>10000 And " & _
                "OpDtl.LCode=LedMst.LCode And LedMst.HCode=12 Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                TxtTotalD.Text = "0.00"
                Do While RsQry.EOF = False
                    .TextMatrix(.Row, 0) = RsQry.Fields("LName")
                    .TextMatrix(.Row, 1) = RsQry.Fields("RTotal")
                    .TextMatrix(.Row, 2) = RsQry.Fields("LCode")
                    TxtTotalD.Text = Val(TxtTotalD.Text) + RsQry.Fields("RTotal")
                    RsQry.MoveNext
                    If RsQry.EOF = False Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                Loop
            Case 7  '   Provincial Grants under Recurring Receipts
                Set RsQry = Nothing
                RsQry.Open "Select EntMst.ECode,EntMst.EName,RpDtl.Amt As RTotal From RpDtl,EntMst Where RpDtl.AcCode=" & mAcCode & "And RpDtl.SrN<9999 And " & _
                "RpDtl.Side='C' And RpDtl.ECode=EntMst.ECode And EntMst.LCode=73", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                TxtTotalD.Text = "0.00"
                Do While RsQry.EOF = False
                    .TextMatrix(.Row, 0) = RsQry.Fields("EName")
                    .TextMatrix(.Row, 1) = RsQry.Fields("RTotal")
                    .TextMatrix(.Row, 2) = RsQry.Fields("ECode")
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

Private Sub PrintRec()
Dim RsRep As New ADODB.Recordset
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "School Memo does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
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
SMemoPrint
SMemoNotePrint
'SMemoFEPrint
End Sub

Private Sub SMemoPrint()
Dim i As Double
Dim RsRep As New ADODB.Recordset
RsRep.Open "Select RepDtl.*,AdtMst.AdtName From RepDtl,AdtMst Where AcCode=" & mAcCode & " And RepDtl.AdtNo=AdtMst.AdtNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpSMPrn"
With VsfHelp1
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Insert Into TmpSMPrn (LName,LAmt,LNote,SrN) Values ('" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & _
        "," & Val(.TextMatrix(i, 1)) & "," & i & ")"
    Next
End With
With VsfHelp
    For i = 1 To .Rows - 1
        If i < VsfHelp1.Rows Then
            DbDataDB.Execute "Update TmpSMPrn Set RName='" & .TextMatrix(i, 0) & "',RAmt=" & Val(.TextMatrix(i, 2)) & ",RNote=" & _
            Val(.TextMatrix(i, 1)) & " Where SrN=" & i
        Else
            DbDataDB.Execute "Insert Into TmpSMPrn (RName,RAmt,RNote,SrN) Values ('" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & _
            "," & Val(.TextMatrix(i, 1)) & "," & i & ")"
        End If
    Next
End With
DbDataDB.CommitTrans
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\ScmRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Formulas(8) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(9) = "mTitle1='Memo of Receipts and Expenditure for the year ended on " & RsRep.Fields("RtDt") & "'"
    .Formulas(10) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(11) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
    .Formulas(12) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(13) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(14) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(15) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(16) = "mTotal='" & TxtCTotal.Text & "'"
    .Formulas(17) = "mCYear='" & Year(RsRep.Fields("RtDt")) & "'"
    .Formulas(18) = "mOYear='" & Year(RsRep.Fields("RfDt")) & "'"
    .Action = 1
    For i = 0 To 18
        .Formulas(i) = ""
    Next
End With
End Sub
Private Sub SMemoNotePrint()
Dim i As Double
Dim RsQ As New ADODB.Recordset
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpNotePrn"
i = 3
With VsfHelp1
    Do While i <= .Rows - 1
        If i = 3 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select LedMst.LName,OpDtl.OpBal As RTotal From OpDtl,LedMst Where OpDtl.AcCode=" & mAcCode & " And OpDtl.LCode<>10000 And OpDtl.LCode=LedMst.LCode And LedMst.HCode=12", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & Space(1) & .TextMatrix(i - 2, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 7 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & mAcCode & _
            " And TmpRpDtl.ECode=EntMst.ECode And EntMst.LCode=73", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & _
                    "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 8 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & mAcCode & _
            " And TmpRpDtl.ECode=EntMst.ECode And EntMst.LCode=78", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & _
                    "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 5 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.ECode=EntMst.ECode And " & _
            "EntMst.LCode=79", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & _
                    "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 10 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.ECode=EntMst.ECode And " & _
            "TmpRpDtl.HCode In (42,44,45,46,47,48,53) And TmpRpDtl.LCode Not In (78,79,80,81,82)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & _
                    "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 11 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.ECode=EntMst.ECode And " & _
            "EntMst.LCode=80", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & _
                    "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 12 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.ECode=EntMst.ECode And " & _
            "EntMst.LCode=81", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & _
                    "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 16 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.ECode=EntMst.ECode And " & _
            "EntMst.LCode In (74,83)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & _
                    "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 17 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.ECode=EntMst.ECode And " & _
            "TmpRpDtl.HCode=14", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & _
                    "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 18 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.ECode=EntMst.ECode And " & _
            "TmpRpDtl.HCode In (10,11)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
                    IIf(RsQ.Fields("RTotal") > 0, " ", "Less: ") + RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 19 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(Side='C',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & mAcCode & _
            " And TmpRpDtl.ECode=EntMst.ECode And (TmpRpDtl.HCode=50 or (TmpRpDtl.LCode=82 or TmpRpDtl.LCode=72))", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 20 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select * From QSMNote where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,SAmt,HCode) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & ".','" & _
                RsQ.Fields("LName") & "'," & RsQ.Fields("Cr") & "," & RsQ.Fields("Dr") & "," & RsQ.Fields("Cr") - RsQ.Fields("Dr") & ",-1)"
                RsQ.MoveNext
            Loop
        End If
        i = i + 1
    Loop
End With
i = 2
With VsfHelp
    Do While i <= .Rows - 1
        If i = 2 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And EntMst.LCode=61", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
                    RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 3 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And EntMst.LCode=60", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 4 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And (TmpRpDtl.ECode In (485,999) or TmpRpDtl.LCode=43 or TmpRpDtl.HCode=19)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 5 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.LCode=63", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 6 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.ECode In (484,549)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 7 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.HCode=8 And TmpRpDtl.LCode<>14", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 8 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.HCode=17", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 9 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.HSide='D' And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.ECode Not In (484,485,549,999) And TmpRpDtl.LCode Not In (5,14,43,60,61,62,63) And TmpRpDtl.HCode Not In (2,7,8,10,11,14,17,19,24,26,28,29,30,31,32,33)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 13 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And (TmpRpDtl.HCode In (24,26,28,29,30,31,32,33) Or TmpRpDtl.LCode=62) And TmpRpDtl.ECode<>999", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 14 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.HCode In (7,2)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 15 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.LCode=14", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 16 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select EntMst.EName As LName,IIF(TmpRpDtl.Side='D',TmpRpDtl.Amt,TmpRpDtl.Amt*-1) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode=" & _
            mAcCode & " And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.HCode=14", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        ElseIf i = 21 And Val(.TextMatrix(i, 1)) <> 0 Then
            Set RsQ = Nothing
            RsQ.Open "Select LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & _
            " And TmpTrialBal.LCode<>10000 And TmpTrialBal.HCode=12 And TmpTrialBal.LCode=LedMst.LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            Do While RsQ.EOF = False
                If RsQ.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                    .TextMatrix(i, 0) & Space(1) & .TextMatrix(i - 2, 0) & "','" & RsQ.Fields("LName") & "'," & RsQ.Fields("RTotal") & ")"
                End If
                RsQ.MoveNext
            Loop
        End If
        i = i + 1
    Loop
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
    .Formulas(9) = "mTitle1='Notes to Memo of Receipts and Expenditure'"
    .Formulas(10) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(11) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(12) = "mTrack='0'"
    .Formulas(13) = "mType=''"
    .Formulas(14) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Action = 1
    For i = 0 To 14
        .Formulas(i) = ""
    Next
End With
End Sub

'Private Sub SMemoFEPrint()
'Dim RsQ As New ADODB.Recordset
'Dim RsRep As New ADODB.Recordset
'Dim i As Integer
'RsRep.Open "Select RepDtl.*,AdtMst.AdtName From RepDtl,AdtMst Where AcCode=" & mAcCode & " And RepDtl.AdtNo=AdtMst.AdtNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
'DbDataDB.BeginTrans
'DbDataDB.Execute "Delete From TmpNotePrn"
'RsQ.Open "Select * From QFurEquip Where AcCode=" & mAcCode & " Order By LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
'Do While RsQ.EOF = False
'    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,LName,LAmt,RAmt,DAmt,CAmt,HCode,SAmt) Values (1,'" & RsQ.Fields("LName") & "'," & IIf(RsQ.Fields("OpDr") = 0, "Null", RsQ.Fields("OpDr")) & "," & _
'    IIf(RsQ.Fields("ADr") = 0, "Null", RsQ.Fields("ADr")) & "," & IIf(RsQ.Fields("ACr") = 0, "Null", RsQ.Fields("ACr")) & "," & IIf((RsQ.Fields("OpDr") + RsQ.Fields("ADr")) - RsQ.Fields("ACr") = 0, "Null", (RsQ.Fields("OpDr") + RsQ.Fields("ADr")) - RsQ.Fields("ACr")) & ",8," & "Null" & ")"
'    RsQ.MoveNext
'Loop
'DbDataDB.CommitTrans
'With RepPrint
'    .Connect = MSCONNECT
'    .ReportFileName = App.Path + "\Report\NteRpt.Rpt"
'    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
'    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
'    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
'    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
'    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
'    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
'    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
'    .Formulas(9) = "mTitle1='Statement of Furniture and Equipment for the year ended on " & RsRep.Fields("RtDt") & "'"
'    .Formulas(10) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
'    .Formulas(11) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
'    .Formulas(12) = "mTrack='0'"
'    .Formulas(13) = "mType=''"
'    .Formulas(14) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
'    .Action = 1
'    For i = 0 To 14
'        .Formulas(i) = ""
'    Next
'End With
'End Sub
