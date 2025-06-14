VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmPlAcCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income And Expenditure Account [Consolidated]"
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
   Icon            =   "FPlAcConTrn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   15690
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
      TabIndex        =   28
      ToolTipText     =   "Exit"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exit"
      Top             =   240
      Width           =   1584
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
      TabIndex        =   26
      ToolTipText     =   "Print"
      Top             =   240
      Width           =   732
   End
   Begin VB.Frame FraClientHelp 
      Height          =   4212
      Left            =   18000
      TabIndex        =   5
      Top             =   1920
      Width           =   14750
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   3852
         Left            =   120
         TabIndex        =   6
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
         FormatString    =   $"FPlAcConTrn.frx":0442
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
      TabIndex        =   3
      Top             =   0
      Width           =   15492
      Begin VB.Frame FraDetail 
         Height          =   6132
         Left            =   18000
         TabIndex        =   22
         Top             =   720
         Width           =   5772
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
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   5640
            Width           =   1300
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
            Picture         =   "FPlAcConTrn.frx":048B
            TabIndex        =   23
            ToolTipText     =   "Cancel"
            Top             =   5640
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfDetail 
            Height          =   5412
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   5580
            _cx             =   9842
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
            FormatString    =   $"FPlAcConTrn.frx":08CD
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
            Width           =   5772
         End
      End
      Begin VB.Frame FraNote 
         Height          =   6132
         Left            =   18000
         TabIndex        =   18
         Top             =   480
         Width           =   6756
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
            Picture         =   "FPlAcConTrn.frx":0916
            TabIndex        =   20
            ToolTipText     =   "Cancel"
            Top             =   5640
            Width           =   975
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
            TabIndex        =   19
            Top             =   5640
            Width           =   1600
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfNote 
            Height          =   5412
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   6540
            _cx             =   11536
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
            FormatString    =   $"FPlAcConTrn.frx":0D58
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
            Index           =   2
            Left            =   0
            Top             =   120
            Width           =   6732
         End
      End
      Begin VB.Frame FraList 
         Height          =   3972
         Left            =   36000
         TabIndex        =   14
         Top             =   1560
         Width           =   7572
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
            Picture         =   "FPlAcConTrn.frx":0DA1
            TabIndex        =   16
            ToolTipText     =   "Cancel"
            Top             =   3500
            Width           =   975
         End
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
            TabIndex        =   15
            Top             =   3500
            Width           =   1776
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfList 
            Height          =   3252
            Left            =   120
            TabIndex        =   17
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
            FormatString    =   $"FPlAcConTrn.frx":11E3
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
         TabIndex        =   13
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   732
      End
      Begin VB.Frame FraMainExp 
         BackColor       =   &H00F7E0FE&
         Height          =   5412
         Left            =   18000
         TabIndex        =   9
         Top             =   4000
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
            TabIndex        =   10
            ToolTipText     =   "Cancel"
            Top             =   4920
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfMainExport 
            Height          =   5052
            Left            =   0
            TabIndex        =   11
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
            FormatString    =   $"FPlAcConTrn.frx":122C
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   9000
         Width           =   1776
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
         Picture         =   "FPlAcConTrn.frx":1275
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
         TabIndex        =   2
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
            Size            =   9.75
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
         FormatString    =   $"FPlAcConTrn.frx":18AB
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
         TabIndex        =   12
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
         TabIndex        =   4
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
Attribute VB_Name = "FrmPlAcCon"
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
Dim RsRep As New ADODB.Recordset
Private Sub CmdCancelD_Click()
    FraDetail.Left = 18000
    FraNote.Left = 180
    VsfNote.SetFocus
End Sub

Private Sub CmdSave_Click()
Dim i As Integer
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete NtDtl.* From NtDtl,HedMst Where NtDtl.RType=1 And NtDtl.AcCode=" & mAcCode & " And NtDtl.HCode=HedMst.HCode And HedMst.HType=0"
With VsfHelp
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) <> 0 Then
            DbDataDB.Execute "Insert Into NtDtl (AcCode,HCode,[Note],RSide,RType) Values (" & mAcCode & "," & Val(.TextMatrix(i, 6)) & "," & Val(.TextMatrix(i, 1)) & ",'D',1)"
        End If
    Next
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 4)) <> 0 Then
            DbDataDB.Execute "Insert Into NtDtl (AcCode,HCode,[Note],RSide,RType) Values (" & mAcCode & "," & Val(.TextMatrix(i, 7)) & "," & Val(.TextMatrix(i, 4)) & ",'C',1)"
        End If
    Next
End With
'DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'INEXP_CON','SAVE_NOTES','" & Date & "','" & Time & "')"
DbDataDB.CommitTrans
MsgBox "Record Saved Successfully.", vbInformation, "Alert"
Exit Sub
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
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
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 And " & _
"AcMst.AcType=1 And AcMst.AcType=GrpMst.GCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.City,AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,'')," & _
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
    SetData
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'INEXP_CON','VIEW','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    VsfHelp.SetFocus
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Cancel"
        FraNote.Left = 26000
    Case "Print"
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'INEXP_CON','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
    Case "Exit"
        DBWorkTmp.Close
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(2).Enabled = True
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
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'INEXP_CON','EXPORT_XLS','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
End Sub
Private Sub ExpMainData()
Dim i As Double
With VsfMainExport
    .Cols = 6
    .Rows = 1
    .TextMatrix(0, 3) = "INCOME"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "NOTE"
    .ColWidth(1) = 700
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1700
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 0) = "EXPENDITURE"
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
VsfMainExport.SaveGrid Environ("USERPROFILE") & "\Desktop\INEXP.XLS", flexFileTabText
MsgBox "Successfully Excel File Generated In " + Environ("USERPROFILE") + "\Desktop\INEXP.XLS", vbInformation, "Alert"
End Sub
Private Sub SetData()
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
    If LsvClient.TextMatrix(LsvClient.Row, 4) <> "Trust" Then
        RsQry.Open "Select ECode,EName,Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where Side='D' And HType=0 And LCode<>107 Group By ECode,EName Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQry.EOF = False
            .TextMatrix(.Row, 0) = RsQry.Fields("EName")
            .TextMatrix(.Row, 2) = RsQry.Fields("Amount")
            .TextMatrix(.Row, 6) = RsQry.Fields("ECode")
            RsQry.MoveNext
            If RsQry.EOF = False Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        Loop
        .Rows = .Rows + 1
        .Row = 1
        Set RsQry = Nothing
        RsQry.Open "Select ECode,EName,Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl where Side='C' And HType=0 And LCode<>108 Group By ECode,EName Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQry.EOF = False
            .TextMatrix(.Row, 3) = RsQry.Fields("EName")
            .TextMatrix(.Row, 5) = RsQry.Fields("Amount")
            .TextMatrix(.Row, 7) = RsQry.Fields("ECode")
            RsQry.MoveNext
            If RsQry.EOF = False Then
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                Else
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                End If
             Else
                .Rows = .Rows + 1
            End If
        Loop
    Else
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
        RsQry.Open "Select HCode,Sum(DBal) AS Amt From TmpTrialBal Where HSide='D' And LCode<>107 Group By HCode Order By HCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQry.EOF = False
            i = .FindRow(RsQry.Fields("HCode"), 1, 6)
            If i > 0 Then .TextMatrix(i, 2) = RsQry.Fields("Amt")
            RsQry.MoveNext
        Loop
        Set RsQry = Nothing
        RsQry.Open "Select HCode,Sum(CBal) AS Amt From TmpTrialBal Where HSide='C' And LCode<>108 Group By HCode Order By HCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQry.EOF = False
            i = .FindRow(RsQry.Fields("HCode"), 1, 7)
            If i > 0 Then .TextMatrix(i, 5) = RsQry.Fields("Amt")
            RsQry.MoveNext
        Loop
    End If
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

Private Sub VsfHelp_EnterCell()
If VsfHelp.Col = 1 Or VsfHelp.Col = 4 Then
    VsfHelp.Editable = flexEDKbd
    VsfHelp.AutoSearch = flexSearchNone
    On Error Resume Next
    SendKeys "{F2}"
Else
    VsfHelp.Editable = flexEDNone
    VsfHelp.AutoSearch = flexSearchFromCursor
End If
End Sub
Private Sub CmdLClose_Click()
    FraList.Left = 36000
End Sub

Private Sub VsfHelp_DblClick()
    FraList.Left = 3000
    SetCmpTotal
    VsfList.SetFocus
End Sub
Private Sub SetCmpTotal()
Dim RsTotal As New ADODB.Recordset
Dim mParentcd As Double
Dim mAmt As Double
Dim i As Integer
Dim mGrp As Double
If InStr(1, mAcList, ",") > 0 Then mParentcd = Mid(mAcList, 1, InStr(1, mAcList, ",") - 1) Else mParentcd = mAcList
If VsfHelp.Col <= 2 Then
    mGrp = Val(VsfHelp.TextMatrix(VsfHelp.Row, 6))
    RsTotal.Open "Select AcCode,IIF(IsNull(Sum(DBal))=False,Sum(DBal),0) As RTotal From TmpTrialBal Where HCode=" & mGrp & " And LCode Not In (107,108) Group By AcCode Order By AcCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
Else
    mGrp = Val(VsfHelp.TextMatrix(VsfHelp.Row, 7))
    RsTotal.Open "Select AcCode,IIF(IsNull(Sum(CBal))=False,Sum(CBal),0) As RTotal From TmpTrialBal Where HCode=" & mGrp & " And LCode Not In (107,108) Group By AcCode Order By AcCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
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
    TxtLTotal.Text = "0.00"
    For i = 1 To .Rows - 1
        TxtLTotal.Text = Val(Format(TxtLTotal.Text)) + Val(.TextMatrix(i, 2))
    Next
    TxtLTotal.Text = Format(TxtLTotal.Text, "0.00")
End With
End Sub
Private Sub VsfList_DblClick()
    FraList.Left = 36000
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
RsLedger.Open "Select LedMst.LName,LedMst.LCode,HedMst.HCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 " & _
"And LedMst.HType=0 And LedMst.LCode Not In (107,108) And (LedMst.AcCode=0 Or LedMst.AcCode=" & mLCode & ") And LedMst.HCode=" & mHeadCode & _
" And LedMst.HCode=SeqMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mLCode & _
") And SeqMst.HCode=HedMst.HCode Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set RsQry = Nothing
With VsfNote
    .Editable = flexEDNone
    .Cols = 4
    .Rows = 1
    .Row = 0
    .TextMatrix(0, 0) = "LEDGER"
    .ColWidth(0) = 4500
    .TextMatrix(0, 1) = "CUR.YEAR"
    .ColWidth(1) = 1600
    .ColFormat(1) = "0.00"
    .ColAlignment(1) = flexAlignRightCenter
    .FixedAlignment(1) = flexAlignRightCenter
    .ColWidth(2) = 0  '  LEDCODE
    .ColWidth(3) = 0   '  HCODE
    .Col = 0
    .Refresh
    .Rows = 2
    .Row = 1
    Do While RsLedger.EOF = False
        .TextMatrix(.Row, 0) = RsLedger.Fields("LName")
        .TextMatrix(.Row, 2) = RsLedger.Fields("LCode")
        .TextMatrix(.Row, 3) = RsLedger.Fields("HCode")
        RsLedger.MoveNext
        If RsLedger.EOF = False Then
            .Rows = .Rows + 1
            .Row = .Row + 1
        End If
    Loop
    For mHeadCode = 1 To .Rows - 1
        Set RsLedger = Nothing
        RsLedger.Open "Select * From TmpTrialBal Where AcCode=" & mLCode & " And LCode Not In (107,108) And LCode=" & Val(.TextMatrix(mHeadCode, 2)), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsLedger.EOF = False Then
            If mSide > 2 Then
                .TextMatrix(mHeadCode, 1) = RsLedger.Fields("ACr") - RsLedger.Fields("ADr")
            Else
                .TextMatrix(mHeadCode, 1) = RsLedger.Fields("ADr") - RsLedger.Fields("ACr")
            End If
        End If
    Next
    mRow = SetProfitAll(mAcCode)
    If mRow > 0 Then
        mLCode = .FindRow(41, 1, 3) 'finding row
        If mLCode > 0 Then .TextMatrix(mLCode, 1) = mRow   'Surplus
    Else
        mLCode = .FindRow(55, 1, 3) 'finding row
        If mLCode > 0 Then .TextMatrix(mLCode, 1) = Abs(mRow)   'Deficit
    End If
End With
SetTotal
End Sub
Private Function ShowLedOpen()
Dim i As Double
Dim RsQ As New ADODB.Recordset
Dim mLCode As Double
mLCode = Val(VsfList.TextMatrix(VsfList.Row, 3))
On Error GoTo XErr
RsQ.Open "Select LCode,OpBal From OpDtl Where AcCode=" & mLCode & " Order By LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With VsfNote
    Do While RsQ.EOF = False
        i = .FindRow(RsQ.Fields("LCode"), 1, 4)
        If i > 0 Then .TextMatrix(i, 1) = RsQ.Fields("OpBal")
        RsQ.MoveNext
    Loop
    SetTotal
End With
Exit Function
XErr:
MsgBox Err.Description
End Function
Private Sub SetTotal()
Dim i As Double
TxtOTotal.Text = "0.00"
With VsfNote
    For i = 1 To .Rows - 1
        TxtOTotal.Text = Val(TxtOTotal.Text) + Val(.TextMatrix(i, 1))
    Next
End With
TxtOTotal.Text = Format(TxtOTotal.Text, "0.00")
End Sub

Private Sub VsfNote_DblClick()
'If VsfNote.Col = 2 Then
    FraDetail.Left = 120
    FraNote.Left = 18000
    SetDetail
    ShowDetail
    VsfDetail.SetFocus
'End If
End Sub
Private Sub SetDetail()
With VsfDetail
    .Rows = 1
    .Cols = 4
    .TextMatrix(0, 0) = "ACTION"
    .ColWidth(0) = 800
    .TextMatrix(0, 1) = "PARTICULAR"
    .ColWidth(1) = 2700
    .TextMatrix(0, 2) = "AMOUNT RS."
    .ColWidth(2) = 1300
    .ColAlignment(2) = flexAlignRightCenter
    .ColFormat(2) = "0.00"
    .ColWidth(3) = 100
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
RsQ.Open "Select IIF(Side=HSide,'Add','Less') As Side,EName,Amt,ECode From TmpCtDtl Where AcCode=" & mLCode & " And LCode=" & VsfNote.TextMatrix(VsfNote.Row, 2) & " Order By Side", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
PlAcPrint
If LsvClient.TextMatrix(LsvClient.Row, 4) = "Trust" Then NotePrint
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
With RepPrint
    .Connect = MSCONNECT
        If LsvClient.TextMatrix(LsvClient.Row, 4) = "Trust" Then
        .ReportFileName = App.Path + "\Report\AIERpt.Rpt"
        .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
        .Formulas(1) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
        .Formulas(2) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
        .Formulas(3) = "mTitle1='Income and Expenditure Account for the year ended on " & RsRep.Fields("RtDt") & "'"
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
        .Formulas(20) = "mActTitle='" & IIf(Len(RsState.Fields("IETitle")) > 0, RsState.Fields("IETitle"), "") & "'"
        .Formulas(21) = "mActSub='" & IIf(Len(RsState.Fields("IESub")) > 0, RsState.Fields("IESub"), "") & "'"
        .Action = 1
        For i = 0 To 21
            .Formulas(i) = ""
        Next
    Else
        .ReportFileName = App.Path + "\Report\IEXRpt.Rpt"
        .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
        .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
        .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
        .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
        .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
        .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
        .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
        .Formulas(8) = "mSubHead=''"
        .Formulas(9) = "mTitle1='Income and Expenditure Account for the year ended on " & RsRep.Fields("RtDt") & "'"
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
Private Sub NotePrint()
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
                    mAcName = Space(3) + LsvClient.TextMatrix(LsvClient.Row, 0)
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
                    RsDataO.Open "Select EName,IIF(Side=HSide,Amt,Amt*-1) As Amount From TmpCtDtl Where AcCode=" & mAcCode & " And HCode=" & mGroup & " Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
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
                    mAcName = Space(3) + LsvClient.TextMatrix(LsvClient.Row, 0)
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
                    RsDataO.Open "Select EName,IIF(Side=HSide,Amt,Amt*-1) As Amount From TmpCtDtl Where AcCode=" & mAcCode & " And HCode=" & mGroup & " Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
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
    .Formulas(13) = "mType='I'"
    .Formulas(14) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
    .Action = 1
    For i = 0 To 14
        .Formulas(i) = ""
    Next
End With
mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
SetParent
End Sub
