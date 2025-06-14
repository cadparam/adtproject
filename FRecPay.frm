VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmRecPay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipts And Payments Entry"
   ClientHeight    =   9570
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   16905
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRecPay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   16905
   Begin VB.Frame FraParent 
      Height          =   6765
      Left            =   18000
      TabIndex        =   52
      Top             =   1200
      Width           =   9705
      Begin VSFlex7Ctl.VSFlexGrid VsfParent 
         Height          =   6375
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   9465
         _cx             =   16695
         _cy             =   11245
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
         BackColor       =   16766928
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   128
         BackColorBkg    =   16766928
         BackColorAlternate=   16766928
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
         FormatString    =   $"FRecPay.frx":0442
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
         Height          =   6615
         Index           =   6
         Left            =   0
         Top             =   120
         Width           =   9705
      End
   End
   Begin VB.Frame FraFloOnLed 
      Height          =   5052
      Left            =   18000
      TabIndex        =   37
      Top             =   1320
      Width           =   9252
      Begin VB.CheckBox ChkSpecific 
         Caption         =   "Client Specific"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   6840
         TabIndex        =   46
         Top             =   240
         Width           =   1900
      End
      Begin VB.TextBox TxtFMainHead 
         Height          =   384
         Left            =   1680
         TabIndex        =   42
         Top             =   720
         Width           =   5016
      End
      Begin VB.TextBox TxtFLedName 
         Height          =   384
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Width           =   5016
      End
      Begin VB.TextBox TxtFType 
         Height          =   384
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1200
         Width           =   2856
      End
      Begin VB.CommandButton CmdLedSaveF 
         Caption         =   "&5 Save"
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
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Save"
         Top             =   1200
         Width           =   852
      End
      Begin VB.CommandButton CmdLedCancelF 
         Caption         =   "&6 Close"
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Save"
         Top             =   1200
         Width           =   852
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfLedHelpF 
         Height          =   3252
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   9012
         _cx             =   15896
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   128
         BackColorBkg    =   16777215
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
         FormatString    =   $"FRecPay.frx":048B
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
         Height          =   4932
         Index           =   4
         Left            =   0
         Top             =   120
         Width           =   9252
      End
      Begin VB.Label LblCompany 
         Caption         =   "Ledger Name :"
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
         TabIndex        =   41
         Top             =   240
         Width           =   1464
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
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   696
      End
      Begin VB.Label LblCompany 
         Caption         =   "Main Head :"
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
         TabIndex        =   39
         Top             =   720
         Width           =   1212
      End
   End
   Begin VB.Frame FraLedHelp 
      Height          =   6012
      Left            =   18000
      TabIndex        =   11
      Top             =   600
      Width           =   15612
      Begin VB.CommandButton CmdNew 
         Caption         =   "&3 New Sub Ledger"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Add"
         Top             =   5550
         Width           =   1800
      End
      Begin VB.CommandButton CmdCancel 
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
         Height          =   375
         Left            =   3960
         TabIndex        =   15
         ToolTipText     =   "Add"
         Top             =   5550
         Width           =   800
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         ToolTipText     =   "Add"
         Top             =   5550
         Width           =   732
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfLedHelp 
         Height          =   5292
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   15420
         _cx             =   27199
         _cy             =   9334
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
         FormatString    =   $"FRecPay.frx":04D4
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
         Height          =   5892
         Index           =   3
         Left            =   0
         Top             =   120
         Width           =   15612
      End
   End
   Begin VB.Frame FraClientHelp 
      Height          =   7965
      Left            =   18000
      TabIndex        =   7
      Top             =   720
      Width           =   14750
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   7575
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   14505
         _cx             =   25585
         _cy             =   13361
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
         FormatString    =   $"FRecPay.frx":051D
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
         Height          =   7815
         Index           =   1
         Left            =   0
         Top             =   120
         Width           =   14745
      End
   End
   Begin VB.Frame FraMst 
      Height          =   9495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   16695
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   1450
      End
      Begin VB.CommandButton CmdRow 
         Caption         =   "New Row"
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Exit"
         Top             =   8925
         Width           =   1100
      End
      Begin VB.TextBox TxtDTotal 
         Alignment       =   1  'Right Justify
         Height          =   372
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   8895
         Width           =   1600
      End
      Begin VB.CommandButton CmdAddLedger 
         Caption         =   "Add Closing Balance Ledger"
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
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Exit"
         Top             =   8925
         Width           =   3000
      End
      Begin VB.Frame FraMainExp 
         BackColor       =   &H00F7E0FE&
         Height          =   5412
         Left            =   18000
         TabIndex        =   43
         Top             =   720
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
            TabIndex        =   44
            ToolTipText     =   "Cancel"
            Top             =   4920
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfMainExport 
            Height          =   5052
            Left            =   0
            TabIndex        =   45
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
            FormatString    =   $"FRecPay.frx":0566
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
      Begin VB.Frame FraFlyOnSubLed 
         Height          =   5052
         Left            =   18000
         TabIndex        =   20
         Top             =   960
         Width           =   12612
         Begin VB.CheckBox ChkSpecificE 
            Caption         =   "Client Specific"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   9240
            TabIndex        =   51
            Top             =   480
            Width           =   1900
         End
         Begin VB.CommandButton CmdNewLedF 
            Caption         =   "&4 New Ledger"
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
            Left            =   7440
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Save"
            Top             =   360
            Width           =   1572
         End
         Begin VB.TextBox TxtLedNameF 
            Height          =   384
            Left            =   1920
            TabIndex        =   25
            Top             =   720
            Width           =   5016
         End
         Begin VB.CommandButton CmdCloseF 
            Caption         =   "&2 Close"
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
            Left            =   8040
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Save"
            Top             =   1200
            Width           =   852
         End
         Begin VB.CommandButton CmdSaveF 
            Caption         =   "&1 Save"
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
            Left            =   6960
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Save"
            Top             =   1200
            Width           =   852
         End
         Begin VB.TextBox TxtMainHeadF 
            Height          =   384
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1200
            Width           =   2136
         End
         Begin VB.TextBox TxtTypeF 
            Height          =   384
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1200
            Width           =   2856
         End
         Begin VB.TextBox TxtNameF 
            Height          =   384
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Width           =   5016
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfHeadHelp 
            Height          =   3252
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   12372
            _cx             =   21823
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
            BackColor       =   8438015
            ForeColor       =   -2147483640
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   128
            BackColorBkg    =   12648384
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
            FormatString    =   $"FRecPay.frx":05AF
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
            Caption         =   "Main Head :"
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
            Left            =   3720
            TabIndex        =   31
            Top             =   1200
            Width           =   1092
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
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   576
         End
         Begin VB.Label LblCompany 
            Caption         =   "Under Ledger :"
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
            TabIndex        =   27
            Top             =   720
            Width           =   1380
         End
         Begin VB.Label LblCompany 
            Caption         =   "Sub Ledger Name :"
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
            TabIndex        =   26
            Top             =   240
            Width           =   1740
         End
         Begin VB.Shape ShpMst 
            BorderWidth     =   2
            Height          =   4932
            Index           =   2
            Left            =   0
            Top             =   120
            Width           =   12612
         End
      End
      Begin VB.CommandButton CmdRow1 
         Caption         =   "New Row"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit"
         Top             =   8925
         Width           =   1100
      End
      Begin VB.TextBox TxtCTotal 
         Alignment       =   1  'Right Justify
         Height          =   372
         Left            =   14040
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   9015
         Width           =   1600
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
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   732
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
         Left            =   10170
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancel"
         Top             =   240
         Width           =   852
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
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   600
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
         Left            =   840
         Picture         =   "FRecPay.frx":05F8
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Add"
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox TxtName 
         Height          =   384
         Left            =   1248
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   5496
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   8205
         Left            =   8400
         TabIndex        =   18
         Top             =   720
         Width           =   8100
         _cx             =   14287
         _cy             =   14473
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
         FormatString    =   $"FRecPay.frx":0C2E
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
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp1 
         Height          =   8205
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   8100
         _cx             =   14287
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
         FormatString    =   $"FRecPay.frx":0C77
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
         Left            =   9360
         TabIndex        =   17
         Top             =   9015
         Width           =   60
      End
      Begin VB.Label LblCompany 
         Caption         =   "Client :"
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
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   708
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   9375
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   16695
      End
   End
   Begin Crystal.CrystalReport RepPrint 
      Left            =   120
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
Attribute VB_Name = "FrmRecPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mAcCode As Double
Dim mAuto As Double
Dim mLedCode As Double
Dim mActive As String
Dim mTrack As String
Dim mEAct As String
Dim DBWorkTmp As New ADODB.Connection
Dim RsQ As New ADODB.Recordset
Private Sub CmdAddLedger_Click()
Dim i As Double
Set RsQ = Nothing
With VsfHelp
    RsQ.Open "Select LName,LCode From HedMst,LedMst Where LCode Not In (Select LCode From QLedList Where AcCode=" & mAcCode & ") And HedMst.HCode=LedMst.HCode And " & _
    "HedMst.HType=1 And (UCase(Mid(HedMst.HName,1,4))='CASH' Or UCase(Mid(HedMst.HName,1,4))='BANK') And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And LedMst.Active=-1 Union " & _
    "Select LName,LCode From HedMst,LedMst Where LCode Not In (Select LCode From QLedList Where AcCode=" & mAcCode & ") And HedMst.HCode=LedMst.HCode And HedMst.HType=1 " & _
    "And (UCase(Mid(HedMst.HName,1,4))='CASH' Or UCase(Mid(HedMst.HName,1,4))='BANK') And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And LedMst.Active=-1", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .Rows = .Rows + 1
        i = .Rows - 1
        .TextMatrix(i, 0) = "Closing Balance"
        .TextMatrix(i, 1) = RsQ.Fields("LName")
        .TextMatrix(i, 3) = RsQ.Fields("LCode")
        RsQ.MoveNext
    Loop
End With
End Sub
Private Sub CmdCancel_Click()
    FraLedHelp.Left = 18000
    VsfHelp.SetFocus
End Sub
Private Sub CmdCloseF_Click()
    FraFlyOnSubLed.Left = 18000
    FraLedHelp.Left = 180
    TxtNameF.SetFocus
End Sub

Private Sub CmdExport_Click()
    ExpMainData
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'RECPAY','EXPORT_XLS','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
End Sub

Private Sub CmdLedCancelF_Click()
    FraFloOnLed.Left = 18000
    FraFlyOnSubLed.Left = 180
    TxtNameF.SetFocus
End Sub

Private Sub CmdLedSaveF_Click()
Dim RsQry As New ADODB.Recordset
If MsgBox("Are you sure to save?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
    If UCase(Mid(VsfLedHelpF.TextMatrix(VsfLedHelpF.Row, 0), 1, 4)) = "CASH" Then
        RsQry.Open "Select LCode As RCode From LedMst Where LCode>=10000 And Active=0", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQry.EOF = False Then
            mLedCode = RsQry.Fields("RCode")
            mEAct = "E"
        Else
            Set RsQry = Nothing
            RsQry.Open "Select IIF(IsNull(Max(LCode))=True,1,Max(LCode)+1) As RCode From LedMst Where LCode>=10000", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            mLedCode = RsQry.Fields("RCode")
            mEAct = "A"
        End If
    Else
        RsQry.Open "Select LCode As RCode From LedMst Where LCode Between 300 And 9999 And Active=0", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQry.EOF = False Then
            mLedCode = RsQry.Fields("RCode")
            mEAct = "E"
        Else
            Set RsQry = Nothing
            RsQry.Open "Select IIF(IsNull(Max(LCode))=True,1,Max(LCode)+1) As RCode From LedMst Where LCode Between 300 And 9999", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            mLedCode = RsQry.Fields("RCode")
            mEAct = "A"
        End If
    End If
    DbDataDB.BeginTrans
    If mEAct = "E" Then
        DbDataDB.Execute "Update LedMst Set LName='" & TxtFLedName.Text & "', HCode=" & Val(VsfLedHelpF.TextMatrix(VsfLedHelpF.Row, 2)) & ", HType=" & IIf(Mid(VsfLedHelpF.TextMatrix(VsfLedHelpF.Row, 1), 1, 1) = "B", 1, 0) & _
        ", Active=-1, AcCode=" & IIf(ChkSpecific.Value = 1, mAcCode, 0) & " Where LCode=" & mLedCode
    Else
        DbDataDB.Execute "Insert InTo LedMst (LCode,LName,HCode,HTYpe,Active,AcCode) Values (" & mLedCode & ",'" & TxtFLedName.Text & _
        "'," & Val(VsfLedHelpF.TextMatrix(VsfLedHelpF.Row, 2)) & "," & IIf(Mid(VsfLedHelpF.TextMatrix(VsfLedHelpF.Row, 1), 1, 1) = "B", 1, 0) & ",1," & _
        IIf(ChkSpecific.Value = 1, mAcCode, 0) & ")"
    End If
    DbDataDB.CommitTrans
    Set RsQry = Nothing
    RsQry.Open "Select LedMst.LName,HedMst.HName,'Balance Sheet',LedMst.LCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 And LedMst.HType=1 And " & _
    "(LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And LedMst.HCode=SeqMst.HCode And SeqMst.HCode=HedMst.HCode And HedMst.HCode<>12 And SeqMst.GCode In (Select AcType " & _
    "From AcMst Where AcCode=" & mAcCode & ") Union All Select LedMst.LName,HedMst.HName,'Income and Expenditure',LedMst.LCode From LedMst,SeqMst" & _
    ",HedMst Where LedMst.Active=-1 And LedMst.HType=0 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And LedMst.HCode=SeqMst.HCode And " & _
    "SeqMst.HCode=HedMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Set VsfHeadHelp.DataSource = RsQry
    With VsfHeadHelp
        .TextMatrix(0, 0) = "LEDGER NAME"
        .ColWidth(0) = 4000
        .TextMatrix(0, 1) = "HEAD NAME"
        .ColWidth(1) = 4500
        .TextMatrix(0, 2) = "MAIN HEAD"
        .ColWidth(2) = 3500
        .ColWidth(3) = 0    'LCode
        .Refresh
    End With
    FraFloOnLed.Left = 18000
    FraFlyOnSubLed.Left = 640
    VsfHeadHelp.SetFocus
End If
End Sub

Private Sub CmdNew_Click()
    FraLedHelp.Left = 18000
    FraFlyOnSubLed.Left = 640
    TxtNameF.SetFocus
End Sub

Private Sub CmdNewLedF_Click()
    FraFlyOnSubLed.Left = 18000
    TxtFLedName.Text = ""
    FraFloOnLed.Left = 640
    TxtFLedName.Text = ""
    TxtFLedName.SetFocus
End Sub

Private Sub CmdOk_Click()
    FraLedHelp.Left = 18000
    VsfHelp.SetFocus
End Sub
Private Sub CmdRow_Click()
With VsfHelp
    .Rows = .Rows + 1
    .Row = .Rows - 1
End With
End Sub
Private Sub CmdRow1_Click()
With VsfHelp1
    .Rows = .Rows + 1
    .Row = .Rows - 1
End With
End Sub

Private Sub CmdSaveF_Click()
If MsgBox("Are you sure to save?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
    GAuto
    If mTrack = "" Then mTrack = "D"
    DbDataDB.BeginTrans
    If mEAct = "E" Then
        DbDataDB.Execute "Update EntMst Set EName='" & TxtNameF.Text & "', LCode=" & Val(VsfHeadHelp.TextMatrix(VsfHeadHelp.Row, 3)) & ", Active=-1, AcCode=" & IIf(ChkSpecificE.Value = 1, mAcCode, 0) & _
        ", ESide='" & mTrack & "' Where ECode=" & mAuto
    Else
        DbDataDB.Execute "Insert InTo EntMst (ECode,EName,LCode,Active,AcCode,ESide) Values (" & mAuto & ",'" & TxtNameF.Text & "'," & Val(VsfHeadHelp.TextMatrix(VsfHeadHelp.Row, 3)) & _
        ",-1," & IIf(ChkSpecificE.Value = 1, mAcCode, 0) & ",'" & mTrack & "')"
    End If
    DbDataDB.CommitTrans
    TxtNameF.Text = ""
    FraFlyOnSubLed.Left = 18000
    Dim RsQry As New ADODB.Recordset
    RsQry.Open "Select EntMst.EName,LedMst.LName As Ledger,HedMst.HName,'Balance Sheet' As HeadType,EntMst.ECode,EntMst.LCode From EntMst,LedMst,SeqMst,HedMst Where " & _
    "EntMst.LCode=LedMst.LCode And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And LedMst.HCode=SeqMst.HCode And SeqMst.HCode=HedMst.HCode " & _
    "And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") Union All Select EntMst.EName,LedMst.LName,HedMst.HName,'Income And Expenditure'," & _
    "EntMst.ECode,EntMst.LCode From EntMst,LedMst,SeqMst,HedMst Where EntMst.LCode=LedMst.LCode And LedMst.HType=0 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & _
    ") And LedMst.HCode=SeqMst.HCode And SeqMst.HCode=HedMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") Order By EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Set VsfLedHelp.DataSource = RsQry
    With VsfLedHelp
        .TextMatrix(0, 0) = "SUB LEDGER"
        .ColWidth(0) = 3800
        .TextMatrix(0, 1) = "LEDGER"
        .ColWidth(1) = 3800
        .TextMatrix(0, 2) = "HEAD"
        .ColWidth(2) = 2500
        .TextMatrix(0, 3) = "HEAD TYPE"
        .ColWidth(3) = 1100
        .ColWidth(4) = 0    'ECODE
        .ColWidth(5) = 0    'LCode
        .Col = 0
        If .Rows > 1 Then .Row = 1
        .Refresh
    End With
    FraLedHelp.Left = 180
    VsfLedHelp.SetFocus
End If
End Sub

Private Sub CmdSearch_Click()
If TxtName.Text <> "" Then
    If TlbSav(0).Enabled = True Then
        If MsgBox("Save data before exit?", vbInformation + vbYesNo, "Confirmation") = vbYes Then SaveData
    End If
End If
TlbSav(0).Enabled = True
FraClientHelp.Left = 180
LsvClient.SetFocus
End Sub

Private Sub CmpMEClose_Click()
    FraMainExp.Left = 18000
End Sub
Private Sub Form_Load()
    Me.Left = 50
    Me.Top = 50
    DBWorkTmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & App.Path + "\LocalDB.Mdb'"
    DBWorkTmp.Open
    SetCombo
    SetData
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
    .ColWidth(7) = 0 'PACode
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
Dim ObjText As Object
For Each ObjText In Me
    If TypeOf ObjText Is TextBox Then ObjText.Text = ""
Next
DBWorkTmp.BeginTrans
DBWorkTmp.Execute "Delete From TmpTrialBal"
DBWorkTmp.Execute "Delete From TmpRpDtl"
DBWorkTmp.CommitTrans
mAcCode = 0
mAuto = 0
mLedCode = 0
mActive = ""
FraFlyOnSubLed.Left = 18000
FraClientHelp.Left = 18000
FraLedHelp.Left = 18000
FraFloOnLed.Left = 18000
FraParent.Left = 18000
mActive = ""
SetData
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
    SetData
    SetLed "C"
    mTrack = "C"
    ShowData
    SetTotal
    TlbSav(0).Enabled = True
    CheckState
    VsfHelp1.SetFocus
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'RECPAY','VIEW_DATA','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
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
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'RECPAY','SAVE_DATA','" & Date & "','" & Time & "')"
                DbDataDB.CommitTrans
                ClearText
                SetTool True
                FraFlyOnSubLed.Left = 18000
                FraClientHelp.Left = 18000
                FraLedHelp.Left = 18000
                FraFloOnLed.Left = 18000
                CmdSearch.SetFocus
            End If
        End If
    Case "Cancel"
        If MsgBox("Are you sure to Cancel ? ", vbInformation + vbYesNo, "Confirmation") = vbYes Then
            TlbSav(0).Enabled = True
            SetData
            ClearText
            SetCombo
            FraFlyOnSubLed.Left = 18000
            FraClientHelp.Left = 18000
            FraLedHelp.Left = 18000
            FraFloOnLed.Left = 18000
            CmdSearch.SetFocus
        End If
    Case "Print"
        If TxtName.Text <> "" Then
            If TlbSav(0).Enabled = True Then
                If MsgBox("Save data before printing?", vbInformation + vbYesNo, "Confirmation") = vbYes Then SaveData
                DbDataDB.BeginTrans
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'RECPAY','SAVE_DATA','" & Date & "','" & Time & "')"
                DbDataDB.CommitTrans
            End If
        End If
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'RECPAY','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
        ClearText
    Case "Exit"
        If TxtName.Text <> "" Then
            If TlbSav(0).Enabled = True Then
                If MsgBox("Save data before exit?", vbInformation + vbYesNo, "Confirmation") = vbYes Then SaveData
                DbDataDB.BeginTrans
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'RECPAY','SAVE_DATA','" & Date & "','" & Time & "')"
                DbDataDB.CommitTrans
            End If
        End If
        DBWorkTmp.Close
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(0).Enabled = mVal
    TlbSav(1).Enabled = True
    TlbSav(2).Enabled = True
    If mUType <> "A" Then
        CmdNew.Enabled = False
        CmdNewLedF.Enabled = False
        CmdLedSaveF.Enabled = False
        CmdSaveF.Enabled = False
    End If
End Function
Private Sub SetData()
With VsfHelp
    .Cols = 5
    .Rows = 1
    .Rows = 2
    .TextMatrix(0, 0) = "EXPENDITURE"
    .ColWidth(0) = 3000
    .TextMatrix(0, 1) = "UNDER"
    .ColWidth(1) = 3000
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1700
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .ColWidth(3) = 0    'CSIDE
    .ColWidth(4) = 0    'TrfCode
    .Col = 1
    .Row = 1
    .Refresh
End With
With VsfHelp1
    .Cols = 5
    .Rows = 1
    .Rows = 2
    .TextMatrix(0, 0) = "INCOME"
    .ColWidth(0) = 3000
    .TextMatrix(0, 1) = "UNDER"
    .ColWidth(1) = 3000
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1700
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .ColWidth(3) = 0    'DSIDE
    .ColWidth(4) = 0    'TrfCode
    .Col = 1
    .Row = 1
    .Refresh
End With
End Sub
Private Sub SetLed(ByVal mSide As String)
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select EntMst.EName,LedMst.LName As Ledger,HedMst.HName,'Balance Sheet' As HeadType,EntMst.ECode,EntMst.LCode From EntMst,LedMst,SeqMst,HedMst Where " & _
"EntMst.ESide In ('B','" & mSide & "') And EntMst.LCode=LedMst.LCode And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And (EntMst.AcCode=0 Or EntMst.AcCode=" & mAcCode & ") And LedMst.HCode=SeqMst.HCode" & _
" And SeqMst.HCode=HedMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And LedMst.Active=-1 And EntMst.Active=-1 Union All Select EntMst.EName,LedMst.LName,HedMst.HName,'Income And Expenditure'" & _
",EntMst.ECode,EntMst.LCode From EntMst,LedMst,SeqMst,HedMst Where EntMst.ESide In ('B','" & mSide & "') And EntMst.LCode=LedMst.LCode And LedMst.HType=0 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And LedMst.Active=-1 " & _
"And EntMst.Active=-1 And LedMst.HCode=SeqMst.HCode And SeqMst.HCode=HedMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") Order By EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfLedHelp.DataSource = RsQry
With VsfLedHelp
    .TextMatrix(0, 0) = "ENTRY"
    .ColWidth(0) = 4300
    .TextMatrix(0, 1) = "LEDGER"
    .ColWidth(1) = 5700
    .TextMatrix(0, 2) = "HEAD"
    .ColWidth(2) = 3500
    .TextMatrix(0, 3) = "HEAD TYPE"
    .ColWidth(3) = 1500
    .ColWidth(4) = 0    'ECODE
    .ColWidth(5) = 0    'LCode
    .Col = 0
    If .Rows > 1 Then .Row = 1
    .Refresh
End With
Set RsQry = Nothing
RsQry.Open "Select LedMst.LName,HedMst.HName,'Balance Sheet',LedMst.LCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 And LedMst.HType=1 And (LedMst.AcCode=0 Or LedMst.AcCode=" & _
mAcCode & ") And LedMst.HCode=SeqMst.HCode And SeqMst.HCode=HedMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") Union All Select LedMst.LName," & _
"HedMst.HName,'Income and Expenditure',LedMst.LCode From LedMst,SeqMst,HedMst Where LedMst.Active=-1 And LedMst.HType=0 And (LedMst.AcCode=0 Or LedMst.AcCode=" & mAcCode & ") And " & _
"LedMst.HCode=SeqMst.HCode And SeqMst.HCode=HedMst.HCode And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfHeadHelp.DataSource = RsQry
With VsfHeadHelp
    .TextMatrix(0, 0) = "LEDGER NAME"
    .ColWidth(0) = 4000
    .TextMatrix(0, 1) = "HEAD NAME"
    .ColWidth(1) = 4500
    .TextMatrix(0, 2) = "MAIN HEAD"
    .ColWidth(2) = 2500
    .ColWidth(3) = 0    'LCode
    .Refresh
End With
Set RsQry = Nothing
RsQry.Open "Select HedMst.HName,'Balance Sheet' As HeadType,SeqMst.HCode From SeqMst,HedMst Where SeqMst.HCode=HedMst.HCode And SeqMst.GCode In (Select AcType " & _
"From AcMst Where AcCode=" & mAcCode & ") Union All Select HedMst.HName,'Income And Expenditure',SeqMst.HCode From SeqMst,HedMst Where SeqMst.HCode=HedMst.HCode " & _
"And SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfLedHelpF.DataSource = RsQry
With VsfLedHelpF
    .TextMatrix(0, 0) = "GROUP"
    .ColWidth(0) = 5500
    .TextMatrix(0, 1) = "HEAD"
    .ColWidth(1) = 2000
    .ColWidth(2) = 0    'GRPCODE
    .Refresh
End With
End Sub

Private Sub VsfHeadHelp_RowColChange()
With VsfHeadHelp
    TxtLedNameF.Text = .TextMatrix(.Row, 0)
    TxtTypeF.Text = .TextMatrix(.Row, 1)
    TxtMainHeadF.Text = .TextMatrix(.Row, 2)
End With
End Sub

Private Sub VsfHelp_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SetTotal
End Sub

Private Sub VsfHelp_Click()
    mTrack = "D"
    SetLed "D"
End Sub

Private Sub VsfHelp_DblClick()
If TxtName.Text <> "" Then
    With VsfHelp
        If .Row = 0 Then
            Exit Sub
        ElseIf Mid(.TextMatrix(.Row, 0), 1, 4) = "Open" Or Mid(.TextMatrix(.Row, 0), 1, 4) = "Clos" Then
            Exit Sub
        End If
        If .Col = 0 Then
            mActive = .Name
            .Editable = flexEDNone
            If VsfLedHelp.Rows > 1 Then
                FraLedHelp.Left = 180
                VsfLedHelp.Col = 0
                VsfLedHelp.Row = 1
                VsfLedHelp.SetFocus
            End If
        ElseIf .Col > 1 Then
            .Editable = flexEDKbd
        End If
    End With
    SetTotal
End If
End Sub

Private Sub VsfHelp_EnterCell()
If TxtName.Text <> "" Then
    mActive = VsfHelp.Name
    With VsfHelp
        If (.Row >= 1 And .Row <= .Rows - 1) And .Col = 2 And Mid(.TextMatrix(.Row, 0), 1, 4) <> "Clos" Then
            .Editable = flexEDKbd
            On Error Resume Next
            SendKeys "{F2}"
        End If
    End With
End If
End Sub

Private Sub VsfHelp_GotFocus()
    mTrack = "D"
    SetLed "D"
End Sub

Private Sub VsfHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    VsfHelp.ToolTipText = VsfHelp.TextMatrix(VsfHelp.Row, 1)
End Sub
Private Sub VsfHelp1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    VsfHelp1.ToolTipText = VsfHelp1.TextMatrix(VsfHelp1.Row, 1)
End Sub

Private Sub VsfHelp1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SetTotal
End Sub

Private Sub VsfHelp1_Click()
    mTrack = "C"
    SetLed "C"
End Sub

Private Sub VsfHelp1_DblClick()
If TxtName.Text <> "" Then
    With VsfHelp1
        If .Row = 0 Then
            Exit Sub
        ElseIf Mid(.TextMatrix(.Row, 0), 1, 4) = "Open" Or Mid(.TextMatrix(.Row, 0), 1, 4) = "Clos" Then
            Exit Sub
        End If
        If .Col = 0 Then
            mActive = .Name
            .Editable = flexEDNone
            If VsfLedHelp.Rows > 1 Then
                FraLedHelp.Left = 180
                VsfLedHelp.Col = 0
                VsfLedHelp.Row = 1
                VsfLedHelp.SetFocus
            End If
        ElseIf .Col > 1 Then
            .Editable = flexEDKbd
        End If
    End With
    SetTotal
End If
End Sub
Private Sub VsfHelp1_EnterCell()
If TxtName.Text <> "" Then
    With VsfHelp1
        If (.Row >= 1 And .Row <= .Rows - 1) And .Col = 2 And Mid(.TextMatrix(.Row, 0), 1, 4) <> "Open" Then
            .Editable = flexEDKbd
            On Error Resume Next
            SendKeys "{F2}"
        End If
    End With
End If
End Sub
Private Sub VsfHelp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then VsfHelp.Cell(flexcpText, VsfHelp.Row, 0, VsfHelp.Row, 3) = ""
End Sub

Private Sub VsfHelp1_GotFocus()
    mTrack = "C"
    SetLed "C"
End Sub

Private Sub VsfHelp1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then VsfHelp1.Cell(flexcpText, VsfHelp1.Row, 0, VsfHelp1.Row, 3) = ""
End Sub
Private Sub VsfHelp_RowColChange()
If VsfHelp.Row = 0 Then VsfHelp.Row = 1
If TxtName.Text <> "" Then
    With VsfHelp
        If .Col = 0 Then
            mActive = .Name
            .Editable = flexEDNone
            If Mid(.TextMatrix(.Row, 0), 1, 4) = "Open" Or Mid(.TextMatrix(.Row, 0), 1, 4) = "Clos" Then Exit Sub
            If VsfLedHelp.Rows > 1 Then
                FraLedHelp.Left = 180
                VsfLedHelp.Col = 0
                VsfLedHelp.Row = 1
                VsfLedHelp.SetFocus
            End If
        ElseIf .Col = 1 Then
            .Editable = flexEDNone
            mActive = .Name
            .Editable = flexEDNone
            If Val(.TextMatrix(.Row, 3)) = 19 Or Val(.TextMatrix(.Row, 3)) = 20 And VsfParent.Rows > 1 Then
                FraParent.Left = 1800
                VsfParent.SetFocus
            End If
        ElseIf .Col > 1 Then
            .Editable = flexEDKbd
        End If
    End With
    SetTotal
End If
End Sub
Private Sub VsfHelp1_RowColChange()
If VsfHelp1.Row = 0 Then VsfHelp1.Row = 1
If TxtName.Text <> "" Then
    With VsfHelp1
        If .Col = 0 Then
            mActive = VsfHelp1.Name
            .Editable = flexEDNone
            If Mid(.TextMatrix(.Row, 0), 1, 4) = "Open" Or Mid(.TextMatrix(.Row, 0), 1, 4) = "Clos" Then Exit Sub
            If VsfLedHelp.Rows > 1 Then
                FraLedHelp.Left = 180
                VsfLedHelp.Col = 0
                VsfLedHelp.Row = 1
                VsfLedHelp.SetFocus
            End If
        ElseIf .Col = 1 Then
            .Editable = flexEDNone
            mActive = .Name
            .Editable = flexEDNone
            If Val(.TextMatrix(.Row, 3)) = 19 Or Val(.TextMatrix(.Row, 3)) = 20 And VsfParent.Rows > 1 Then
                FraParent.Left = 1800
                VsfParent.SetFocus
            End If
        ElseIf .Col > 1 Then
            .Editable = flexEDKbd
        End If
    End With
    SetTotal
End If
End Sub
Private Sub VsfLedHelp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If mActive = VsfHelp.Name Then
        With VsfLedHelp
            If VsfHelp.Col <= 1 Then
                VsfHelp.TextMatrix(VsfHelp.Row, 0) = .TextMatrix(.Row, 0)
                VsfHelp.TextMatrix(VsfHelp.Row, 1) = .TextMatrix(.Row, 1)
                VsfHelp.TextMatrix(VsfHelp.Row, 3) = .TextMatrix(.Row, 4)
            End If
        End With
    End If
    If mActive = VsfHelp1.Name Then
        With VsfLedHelp
            If VsfHelp1.Col <= 1 Then
                VsfHelp1.TextMatrix(VsfHelp1.Row, 0) = .TextMatrix(.Row, 0)
                VsfHelp1.TextMatrix(VsfHelp1.Row, 1) = .TextMatrix(.Row, 1)
                VsfHelp1.TextMatrix(VsfHelp1.Row, 3) = .TextMatrix(.Row, 4)
            End If
        End With
    End If
    FraLedHelp.Left = 18000
    If mActive = VsfHelp.Name Then VsfHelp.SetFocus Else VsfHelp1.SetFocus
ElseIf KeyCode = 27 Then
    FraLedHelp.Left = 18000
    If mActive = VsfHelp.Name Then VsfHelp.SetFocus Else VsfHelp1.SetFocus
End If
End Sub
Private Sub SetTotal()
Dim i As Double
TxtCTotal.Text = "0.00"
TxtDTotal.Text = "0.00"
With VsfHelp
    For i = 1 To .Rows - 1
        TxtCTotal.Text = Val(TxtCTotal.Text) + Val(.TextMatrix(i, 2))
    Next
End With
With VsfHelp1
    For i = 1 To .Rows - 1
        TxtDTotal.Text = Val(TxtDTotal.Text) + Val(.TextMatrix(i, 2))
    Next
End With
TxtCTotal.Text = Format(TxtCTotal.Text, "0.00")
TxtDTotal.Text = Format(TxtDTotal.Text, "0.00")
If Val(TxtCTotal.Text) - Val(TxtDTotal.Text) = 0 Then LblTotal.Caption = "" Else LblTotal.Caption = "Dif.Rs.-->" & Format(CStr(Abs(Val(TxtCTotal.Text) - Val(TxtDTotal.Text))), "0.00")
End Sub
Private Function SaveData()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From RpDtl Where AcCode=" & mAcCode
With VsfHelp
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) <> "Closing Balance" Then
            If Val(.TextMatrix(i, 2)) <> 0 Then
                DbDataDB.Execute "Insert InTo RpDtl (AcCode,ECode,Side,Amt,SrN,TrfCode) Values (" & mAcCode & "," & Val(.TextMatrix(i, 3)) & ",'D'," & Val(.TextMatrix(i, 2)) & _
                "," & i & "," & IIf((Val(.TextMatrix(i, 3)) = 20 And Val(.TextMatrix(i, 4)) <> 0), Val(.TextMatrix(i, 4)), 0) & ")"
            End If
        Else
            If Val(.TextMatrix(i, 2)) <> 0 Then
                DbDataDB.Execute "Insert InTo RpDtl (AcCode,ECode,Side,Amt,SrN,TrfCode) Values (" & mAcCode & "," & Val(.TextMatrix(i, 3)) & ",'C'," & Val(.TextMatrix(i, 2)) & _
                "," & i + 9999 & ",0)"
            End If
        End If
    Next
End With
With VsfHelp1
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) <> "Opening Balance" Then
            If Val(.TextMatrix(i, 2)) <> 0 Then
                DbDataDB.Execute "Insert InTo RpDtl (AcCode,ECode,Side,Amt,SrN,TrfCode) Values (" & mAcCode & "," & Val(.TextMatrix(i, 3)) & ",'C'," & Val(.TextMatrix(i, 2)) & _
                "," & i & "," & IIf((Val(.TextMatrix(i, 3)) = 19 And Val(.TextMatrix(i, 4)) <> 0), Val(.TextMatrix(i, 4)), 0) & ")"
            End If
        End If
    Next
End With
DbDataDB.CommitTrans
MsgBox "Record Saved.", vbInformation, "Alert"
Exit Function
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Function
Private Sub ShowData()
Dim i As Double
Dim RsQ1 As New ADODB.Recordset
Dim RsClient As New ADODB.Recordset
Set RsQ = Nothing
With VsfHelp
    .Col = 2
    i = 1
    RsQ.Open "Select EntMst.EName,RpDtl.ECode,RpDtl.Side,RpDtl.Amt,RpDtl.SrN,LedMst.LName,RpDtl.TrfCode From RpDtl,EntMst,LedMst Where RpDtl.AcCode=" & mAcCode & " And RpDtl.Side='D'" & _
    " And RpDtl.ECode=EntMst.ECode And EntMst.LCode=LedMst.LCode Order By EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        If i = .Rows Then .Rows = .Rows + 1
        .TextMatrix(i, 0) = RsQ.Fields("EName")
        .TextMatrix(i, 1) = RsQ.Fields("LName")
        .TextMatrix(i, 2) = RsQ.Fields("Amt")
        .TextMatrix(i, 3) = RsQ.Fields("ECode")
        .TextMatrix(i, 4) = RsQ.Fields("TrfCode")
        If RsQ.Fields("TrfCode") <> 0 Then
            Set RsClient = Nothing
            RsClient.Open "Select * From AcMst Where AcCode=" & RsQ.Fields("TrfCode"), DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsClient.EOF = False Then .TextMatrix(i, 0) = IIf(IsNull(RsClient.Fields("City")) = False, RsClient.Fields("AcName") & ", " & RsClient.Fields("City"), RsClient.Fields("AcName"))
        End If
        RsQ.MoveNext
        i = i + 1
        If RsQ.EOF = True Then
            Exit Do
        ElseIf i = .Rows - 1 Then
            .Row = i
            CmdRow_Click
        End If
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select OpDtl.AcCode,OpDtl.LCode,OpDtl.OpBal,LedMst.LName From OpDtl,HedMst,LedMst Where OpDtl.AcCode=" & mAcCode & " And OpDtl.LCode=LedMst.LCode And " & _
    "HedMst.HCode=LedMst.HCode And HedMst.HType=1 And (UCase(Mid(HedMst.HName,1,4))='CASH' Or UCase(Mid(HedMst.HName,1,4))='BANK') And LedMst.LCode Not In (Select " & _
    "RpDtl.ECode From RpDtl,LedMst Where RpDtl.AcCode=" & mAcCode & " And RpDtl.Side='C' And RpDtl.SrN>=9999 And RpDtl.ECode=LedMst.LCode And LedMst.HCode=12)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    RsQ1.Open "Select LedMst.LName,LedMst.LCode,RpDtl.Side,RpDtl.Amt,RpDtl.SrN,LedMst.LName From RpDtl,LedMst Where RpDtl.AcCode=" & mAcCode & " And RpDtl.Side='C'" & _
    " And RpDtl.SrN>=9999 And RpDtl.ECode=LedMst.LCode Order By LedMst.LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If .TextMatrix(.Rows - 1, 0) = "" Then
        i = .Rows - 1
    Else
        .Rows = .Rows + 1
        i = .Rows - 1
    End If
    If RsQ.EOF = False Then
        Do While RsQ.EOF = False
            .TextMatrix(i, 0) = "Closing Balance"
            .TextMatrix(i, 1) = RsQ.Fields("LName")
            .TextMatrix(i, 3) = RsQ.Fields("LCode")
            If RsQ1.BOF = False Then
                RsQ1.MoveFirst
                RsQ1.Find "LCode=" & RsQ.Fields("LCode")
            End If
            If RsQ1.EOF = False Then
                .TextMatrix(i, 2) = RsQ1.Fields("Amount")
            Else
                If RsQ1.BOF = False Then RsQ1.MoveFirst
            End If
            i = i + 1
            RsQ.MoveNext
            If RsQ.EOF = False Then
                .Rows = .Rows + 1
            End If
        Loop
        Do While RsQ1.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = "Closing Balance"
            .TextMatrix(i, 1) = RsQ1.Fields("LName")
            .TextMatrix(i, 2) = RsQ1.Fields("Amt")
            .TextMatrix(i, 3) = RsQ1.Fields("LCode")
            RsQ1.MoveNext
            If RsQ1.EOF = False Then
                .Rows = .Rows + 1
                i = .Rows - 1
            End If
        Loop
    Else
        Do While RsQ1.EOF = False
            .TextMatrix(i, 0) = "Closing Balance"
            .TextMatrix(i, 1) = RsQ1.Fields("LName")
            .TextMatrix(i, 2) = RsQ1.Fields("Amt")
            .TextMatrix(i, 3) = RsQ1.Fields("LCode")
            RsQ1.MoveNext
            If RsQ1.EOF = False Then
                .Rows = .Rows + 1
                i = .Rows - 1
            End If
        Loop
    End If
End With
With VsfHelp1
    Set RsQ = Nothing
    RsQ.Open "Select OpDtl.AcCode,OpDtl.LCode,OpDtl.OpBal,LedMst.LName From OpDtl,HedMst,LedMst Where OpDtl.AcCode=" & mAcCode & " And OpDtl.LCode=LedMst.LCode And " & _
    "HedMst.HCode=LedMst.HCode And HedMst.HType=1 And (UCase(Mid(HedMst.HName,1,4))='CASH' Or UCase(Mid(HedMst.HName,1,4))='BANK')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    i = 1
    Do While RsQ.EOF = False
        .TextMatrix(i, 0) = "Opening Balance"
        .TextMatrix(i, 1) = RsQ.Fields("LName")
        .TextMatrix(i, 2) = RsQ.Fields("OpBal")
        i = i + 1
        RsQ.MoveNext
        If RsQ.EOF = False Then .Rows = .Rows + 1
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select EntMst.EName,RpDtl.ECode,RpDtl.Side,RpDtl.Amt,RpDtl.SrN,LedMst.LName,RpDtl.TrfCode From RpDtl,EntMst,LedMst Where RpDtl.AcCode=" & mAcCode & " And RpDtl.Side='C'" & _
    " And RpDtl.SrN<9999 And RpDtl.ECode=EntMst.ECode And EntMst.LCode=LedMst.LCode Order By EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        If i = .Rows Then .Rows = .Rows + 1
        .TextMatrix(i, 0) = RsQ.Fields("EName")
        .TextMatrix(i, 1) = RsQ.Fields("LName")
        .TextMatrix(i, 2) = RsQ.Fields("Amt")
        .TextMatrix(i, 3) = RsQ.Fields("ECode")
        .TextMatrix(i, 4) = RsQ.Fields("TrfCode")
        If RsQ.Fields("TrfCode") <> 0 Then
            Set RsClient = Nothing
            RsClient.Open "Select * From AcMst Where AcCode=" & RsQ.Fields("TrfCode"), DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsClient.EOF = False Then .TextMatrix(i, 0) = IIf(IsNull(RsClient.Fields("City")) = False, RsClient.Fields("AcName") & ", " & RsClient.Fields("City"), RsClient.Fields("AcName"))
        End If
        RsQ.MoveNext
        i = i + 1
        If RsQ.EOF = True Then
            Exit Do
        ElseIf i = .Rows - 1 Then
            .Row = i
            CmdRow1_Click
        End If
    Loop
    .Refresh
End With
End Sub
Private Sub GAuto()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select ECode As RCode From EntMst Where Active=0", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQry.EOF = False Then
    mAuto = RsQry.Fields("RCode")
    mEAct = "E"
Else
    Set RsQry = Nothing
    RsQry.Open "Select IIF(IsNull(Max(ECode))=True,1,Max(ECode)+1) As RCode From EntMst", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    mAuto = RsQry.Fields("RCode")
    mEAct = "A"
End If
End Sub
Private Sub VsfLedHelpF_RowColChange()
With VsfLedHelpF
    TxtFMainHead.Text = .TextMatrix(.Row, 0)
    TxtFType.Text = .TextMatrix(.Row, 1)
End With
End Sub
Private Sub ExpMainData()
Dim i As Double
With VsfMainExport
    .Cols = 6
    .Rows = 1
    .TextMatrix(0, 0) = "EXPENDITURE"
    .ColWidth(0) = 3000
    .TextMatrix(0, 1) = "UNDER"
    .ColWidth(1) = 3000
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "INCOME"
    .ColWidth(3) = 3000
    .TextMatrix(0, 4) = "UNDER"
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
End With
VsfMainExport.SaveGrid Environ("USERPROFILE") & "\Desktop\RECPAY.XLS", flexFileTabText
MsgBox "Successfully Excel File Generated In " + Environ("USERPROFILE") + "\Desktop\RECPAY.XLS", vbInformation, "Alert"
End Sub

Private Sub CheckState()
Dim mString As String
Set RsQ = Nothing
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
Private Sub PrintRec()
Dim i As Double
Dim m As Double
Dim mHead As String
Dim mClient1 As String
Dim mTrack As Boolean
Dim RsRep As New ADODB.Recordset
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Receipt Payment Account does not tally. Report cannot be printed.", vbCritical, "Alert"
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
DBWorkTmp.BeginTrans
DBWorkTmp.Execute "Delete From TmpTrialBal"
DBWorkTmp.Execute "Delete From TmpRpDtl"
DBWorkTmp.CommitTrans
DBWorkTmp.BeginTrans
Set RsQ = Nothing
RsQ.Open "Select * From QTrialBal Where AcCode  = " & mAcCode & " And HCode=12", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    DBWorkTmp.Execute "Insert InTo TmpTrialBal (HType,HSide,AcCode,HCode,LCode,OpDr,OpCr,ADr,ACr,DBal,CBal) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & _
    "'," & RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("OpDr") & "," & RsQ.Fields("OpCr") & "," & RsQ.Fields("ADr") & _
    "," & RsQ.Fields("ACr") & "," & RsQ.Fields("DBal") & "," & RsQ.Fields("CBal") & ")"
    RsQ.MoveNext
Loop
Set RsQ = Nothing
RsQ.Open "Select HType,HSide,AcCode,HCode,LCode,ECode,TrfCode,Side,Sum(Amt) As Amt From QRpDtl Where AcCode = " & mAcCode & " Group By AcCode,HType,HSide,Side,HCode,LCode,ECode,TrfCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    DBWorkTmp.Execute "Insert InTo TmpRpDtl (HType,HSide,AcCode,HCode,LCode,ECode,TrfCode,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
    RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & "," & RsQ.Fields("TrfCode") & ",'" & RsQ.Fields("Side") & _
    "'," & RsQ.Fields("Amt") & ")"
    RsQ.MoveNext
Loop
DBWorkTmp.CommitTrans
mTrack = False
SetData
With VsfHelp
    .Col = 2
    i = 1
    Set RsQ = Nothing
    RsQ.Open "Select EntMst.EName,Sum(TmpRpDtl.Amt) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.Side='D' And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.TrfCode=0 Group By EntMst.EName Union All " & _
    " Select IIF(IsNull(AcMst.City)=False,(Mid(EntMst.EName,1,Len(EntMst.EName)-4) + AcMst.AcName +', '+ AcMst.City),(Mid(EntMst.EName,1,Len(EntMst.EName)-4) + AcMst.AcName)),Sum(TmpRpDtl.Amt) As RTotal From TmpRpDtl,EntMst,AcMst " & _
    " Where TmpRpDtl.Side='D' And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.TrfCode=AcMst.AcCode Group By IIF(IsNull(AcMst.City)=False,(Mid(EntMst.EName,1,Len(EntMst.EName)-4) + AcMst.AcName +', '+ AcMst.City),(Mid(EntMst.EName,1,Len(EntMst.EName)-4) + AcMst.AcName)) Order By EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        If i = .Rows Then .Rows = .Rows + 1
        .TextMatrix(i, 0) = RsQ.Fields("EName")
        .TextMatrix(i, 2) = RsQ.Fields("RTotal")
        RsQ.MoveNext
        i = i + 1
        If RsQ.EOF = True Then
            Exit Do
        ElseIf i = .Rows - 1 Then
            .Row = i
            CmdRow_Click
        End If
    Loop
    Set RsQ = Nothing
    
    RsQ.Open "Select IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName) As LName,LedMst.LCode,TmpTrialBal.DBal From TmpTrialBal,LedMst Where TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.HCode=12 Order By IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
'    i = i + 1
    .Rows = .Rows + 1
    Do While RsQ.EOF = False
        .TextMatrix(i, 0) = "Closing Balance"
'        .Rows = .Rows + 1
        If RsQ.Fields("DBal") <> 0 Then
            .TextMatrix(i, 1) = RsQ.Fields("LName")
            .TextMatrix(i, 2) = RsQ.Fields("DBal")
            i = i + 1
        End If
        RsQ.MoveNext
        If RsQ.EOF = False Then .Rows = .Rows + 1
    Loop
End With
With VsfHelp1
    Set RsQ = Nothing
    RsQ.Open "Select IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName) As LName,LedMst.LCode,TmpTrialBal.OpDr From TmpTrialBal,LedMst Where TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.HCode=12 Order By IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    i = 1
    Do While RsQ.EOF = False
    .TextMatrix(i, 0) = "Opening Balance"
        If RsQ.Fields("OpDr") <> 0 Then
            .TextMatrix(i, 1) = RsQ.Fields("LName")
            .TextMatrix(i, 2) = RsQ.Fields("OpDr")
        End If
        i = i + 1
        RsQ.MoveNext
        If RsQ.EOF = False Then .Rows = .Rows + 1
    Loop
    .Rows = .Rows + 1
    Set RsQ = Nothing
    RsQ.Open "Select EntMst.EName,Sum(TmpRpDtl.Amt) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.Side='C' And TmpRpDtl.ECode=EntMst.ECode And TrfCode=0 Group By EntMst.EName Union All " & _
    " Select IIF(IsNull(AcMst.City)=False,(Mid(EntMst.EName,1,Len(EntMst.EName)-4) + AcMst.AcName +', '+ AcMst.City),(Mid(EntMst.EName,1,Len(EntMst.EName)-4) + AcMst.AcName)),Sum(TmpRpDtl.Amt) As RTotal " & _
    " From TmpRpDtl,EntMst,AcMst Where TmpRpDtl.Side='C' And TmpRpDtl.ECode=EntMst.ECode And TrfCode=AcMst.AcCode Group By IIF(IsNull(AcMst.City)=False,(Mid(EntMst.EName,1,Len(EntMst.EName)-4) + AcMst.AcName +', '+ AcMst.City),(Mid(EntMst.EName,1,Len(EntMst.EName)-4) + AcMst.AcName)) " & _
    " Order By EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        If i = .Rows Then .Rows = .Rows + 1
        .TextMatrix(i, 0) = RsQ.Fields("EName")
        .TextMatrix(i, 2) = RsQ.Fields("RTotal")
        RsQ.MoveNext
        i = i + 1
        If RsQ.EOF = True Then
            Exit Do
        ElseIf i = .Rows - 1 Then
            .Row = i
            CmdRow1_Click
        End If
    Loop
    .Refresh
End With
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpRecPrn"
m = 0
With VsfHelp
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) <> "" Then
            If UCase(Mid(.TextMatrix(i, 0), 1, 4)) = "CLOS" Then
                If mTrack = False Then
                    m = m + 1
                    DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN,LAmt) Values ('',0," & m & ",0)"
                    m = m + 1
                    DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN,LAmt) Values ('By Closing Balance',0," & m & ",0)"
                    mTrack = True
                    m = m + 1
                    If (.TextMatrix(i, 2) <> "") Then DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN,LAmt) Values ('" & .TextMatrix(i, 1) & "'," & Val(.TextMatrix(i, 2)) & "," & m & ",0)"
                Else
                    m = m + 1
                    If (.TextMatrix(i, 2) <> "") Then DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN,LAmt) Values ('" & .TextMatrix(i, 1) & "'," & Val(.TextMatrix(i, 2)) & "," & m & ",0)"
                End If
            Else
                m = m + 1
                DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN) Values ('By " & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & "," & m & ")"
            End If
        End If
    Next
End With
With VsfHelp1
    i = 1
    m = 1
    Do While i <= .Rows + 1
        If VsfHelp.Rows > i Then
            If VsfHelp1.Rows = m Then Exit Do
            If .TextMatrix(m, 0) <> "" Then
                If UCase(Mid(.TextMatrix(m, 0), 1, 4)) = "OPEN" Then
                    If i = 1 Then
                        DbDataDB.Execute "Update TmpRecPrn Set LName='To Opening Balance',LAmt=0 Where SrN=" & i
                        i = i + 1
                    End If
                    DbDataDB.Execute "Update TmpRecPrn Set LName='" & .TextMatrix(m, 1) & "',LAmt=" & Val(.TextMatrix(m, 2)) & " Where SrN=" & i
                    i = i + 1
                    m = m + 1
                Else
                    If UCase(Mid(.TextMatrix(m - 1, 0), 1, 4)) = "OPEN" Then
                        DbDataDB.Execute "Update TmpRecPrn Set LName='',LAmt=0 Where SrN=" & i
                        DbDataDB.Execute "Update TmpRecPrn Set LName='To " & .TextMatrix(m, 0) & "',LAmt=" & Val(.TextMatrix(m, 2)) & " Where SrN=" & i + 1
                        m = m + 1
                        i = i + 2
                    Else
                        If VsfHelp.TextMatrix(m, 0) <> "" Then
                            DbDataDB.Execute "Update TmpRecPrn Set LName='To " & .TextMatrix(m, 0) & "',LAmt=" & Val(.TextMatrix(m, 2)) & " Where SrN=" & i '+ 1
                            m = m + 1
                            i = i + 1
                        Else
                            DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('To " & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & "," & i & ")"
                            m = m + 1
                            i = i + 1
                        End If
                    End If
                End If
            Else
                i = i + 1
                m = m + 1
            End If
        Else
            If VsfHelp1.Rows = m Then Exit Do
            If .TextMatrix(m, 0) <> "" Then
                If UCase(Mid(.TextMatrix(m, 0), 1, 4)) = "OPEN" Then
                    DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('" & .TextMatrix(m, 1) & "'," & Val(.TextMatrix(m, 2)) & "," & m & ")"
                    m = m + 1
                    i = i + 1
                Else
                    If UCase(Mid(.TextMatrix(m - 1, 0), 1, 4)) = "OPEN" Then
                        If i = VsfHelp.Rows Then
                            DbDataDB.Execute "Update TmpRecPrn Set LName='',LAmt=0 where SrN=" & i & ""
                        Else
                            DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('',0," & i & ")"
                        End If
                        DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('To " & .TextMatrix(m, 0) & "'," & Val(.TextMatrix(m, 2)) & "," & m + VsfHelp1.Rows + 1 & ")"
                        m = m + 1
                        i = i + 1
                    Else
                        DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('To " & .TextMatrix(m, 0) & "'," & Val(.TextMatrix(m, 2)) & "," & m + VsfHelp1.Rows + 1 & ")"
                        m = m + 1
                        i = i + 1
                    End If
                End If
            Else
                m = m + 1
                i = i + 1
            End If
        End If
    Loop
End With
DbDataDB.CommitTrans
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\RpxRpt.Rpt"
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
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) & ", " & LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
        .Formulas(8) = "mSubHead=''"
        mClient1 = CStr(LsvClient.TextMatrix(LsvClient.Row, 0)) & ", " & CStr(LsvClient.TextMatrix(LsvClient.Row, 2))
    End If
    .Formulas(9) = "mTitle1='Memorandum of Receipts and Payments for the year ended on " & RsRep.Fields("RtDt") & "'"
    .Formulas(10) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(11) = "mDate='Date: " & RsRep.Fields("RPDt") & "'"
    .Formulas(12) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(13) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(14) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(15) = "mClient='" & IIf(Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0, TxtName.Text + IIf(Len(LsvClient.TextMatrix(LsvClient.Row, 2)) > 0, ", " + LsvClient.TextMatrix(LsvClient.Row, 2), ""), "") & "'"
    .Formulas(16) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(17) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(18) = "mLTitle='RECEIPTS'"
    .Formulas(19) = "mRTitle='PAYMENTS'"
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        If Len(LsvClient.TextMatrix(LsvClient.Row, 6)) > 0 Then
            mHead = LsvClient.TextMatrix(LsvClient.Row, 6)
        ElseIf Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0 Then
            mHead = LsvClient.TextMatrix(LsvClient.Row, 3)
        Else
            mHead = LsvClient.TextMatrix(LsvClient.Row, 0) & ", " & LsvClient.TextMatrix(LsvClient.Row, 2)
        End If
    .Formulas(20) = "mSub='This Memorandum of Receipts & Payments is issued for the sole purpose of internal use of " & mHead & " and should not be presented before any third parties/agencies/authorities without our consent.'"
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
 End With
End Sub
Private Sub SetParent()
Dim RsQ As New ADODB.Recordset
Dim mPaCode As Double
Dim mPAType As Integer
Dim mAcList As String
If LsvClient.TextMatrix(LsvClient.Row, 4) = "Trust" And LsvClient.TextMatrix(LsvClient.Row, 7) = 0 Then
    RsQ.Open "Select * From QGroup Where SACode=" & mAcCode & " And AcCode<>" & mAcCode & " And AcType<>2 and PAType<>2 Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
ElseIf LsvClient.TextMatrix(LsvClient.Row, 4) = "FC" Then
    RsQ.Open "Select * From QGroup Where PACode=" & mAcCode & " And AcCode<>" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Else
    RsQ.Open "Select * From QGroup Where AcCode=" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    mPAType = RsQ.Fields("PAType")
    If mPAType <> 2 Then
        mPaCode = CStr(RsQ.Fields("SACode"))
    Else
        mPaCode = CStr(RsQ.Fields("PACode"))
    End If
    Set RsQ = Nothing
    If mPAType <> 2 Then
        RsQ.Open "Select * From QGroup Where (SACode=" & mPaCode & " Or PACode=" & mPaCode & " Or AcCode=" & mPaCode & ") And AcCode<>" & mAcCode & " And AcType<>2 and PAType<>2 Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Else
        RsQ.Open "Select * From QGroup Where (SACode=" & mPaCode & " Or PACode=" & mPaCode & " Or AcCode=" & mPaCode & ") And AcCode<>" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    End If
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
If Len(mAcList) <> 0 Then
RsQ.Open "Select FileNo,AcName,City,AcCode From AcMst Where AcCode In (" & mAcList & ") Order By FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set VsfParent.DataSource = RsQ
With VsfParent
    .TextMatrix(0, 0) = "FILE NO"
    .ColWidth(0) = 1200
    .TextMatrix(0, 1) = "NAME"
    .ColWidth(1) = 4500
    .TextMatrix(0, 2) = "CITY"
    .ColWidth(2) = 1500
    .ColWidth(3) = 0   'ACCODE
    .Refresh
End With
End If
End Sub
Private Sub VsfParent_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If mActive = VsfHelp.Name Then
        With VsfParent
            If VsfHelp.Col <= 1 Then
                VsfHelp.TextMatrix(VsfHelp.Row, 4) = .TextMatrix(.Row, 3)
                VsfHelp.TextMatrix(VsfHelp.Row, 0) = .TextMatrix(.Row, 1)
                End If
        End With
    End If
    If mActive = VsfHelp1.Name Then
        With VsfParent
            If VsfHelp1.Col <= 1 Then
                VsfHelp1.TextMatrix(VsfHelp1.Row, 4) = .TextMatrix(.Row, 3)
                VsfHelp1.TextMatrix(VsfHelp1.Row, 0) = .TextMatrix(.Row, 1)
            End If
        End With
    End If
    FraParent.Left = 18000
    If mActive = VsfHelp.Name Then
        VsfHelp.Col = 2
        VsfHelp.SetFocus
    Else
        VsfHelp1.Col = 2
        VsfHelp1.SetFocus
    End If
ElseIf KeyCode = 27 Then
    FraParent.Left = 18000
    If mActive = VsfHelp.Name Then
        VsfHelp.Col = 2
        VsfHelp.SetFocus
    Else
        VsfHelp1.Col = 2
        VsfHelp1.SetFocus
    End If
End If
End Sub
