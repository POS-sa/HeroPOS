VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProducts 
   Caption         =   "Products"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   14280
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picProducts 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      HasDC           =   0   'False
      Height          =   9465
      Index           =   0
      Left            =   0
      ScaleHeight     =   9465
      ScaleWidth      =   19035
      TabIndex        =   0
      Top             =   0
      Width           =   19035
      Begin VSFlex8Ctl.VSFlexGrid grdProd 
         Bindings        =   "frmProducts.frx":0000
         Height          =   2670
         Left            =   0
         TabIndex        =   7
         Top             =   6210
         Width           =   13755
         _cx             =   24262
         _cy             =   4710
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   14737632
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmProducts.frx":0016
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin MSAdodcLib.Adodc adoData 
            Height          =   375
            Left            =   3780
            Top             =   600
            Visible         =   0   'False
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   661
            ConnectMode     =   1
            CursorLocation  =   2
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "adoData"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
      Begin VB.PictureBox picPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   5160
         ScaleHeight     =   3705
         ScaleWidth      =   8385
         TabIndex        =   2
         Top             =   1740
         Visible         =   0   'False
         Width           =   8415
         Begin VSFlex8Ctl.VSFlexGrid grdPrices 
            Height          =   3720
            Left            =   0
            TabIndex        =   101
            Top             =   0
            Width           =   8400
            _cx             =   14817
            _cy             =   6562
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16642749
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   16645618
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   6
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   1430
            ColWidthMax     =   1430
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmProducts.frx":008E
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   5
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
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
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picDepartments 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3825
         Left            =   5220
         ScaleHeight     =   3825
         ScaleWidth      =   8415
         TabIndex        =   3
         Top             =   1740
         Visible         =   0   'False
         Width           =   8415
         Begin VSFlex8Ctl.VSFlexGrid grdMinor1 
            Height          =   3810
            Left            =   5310
            TabIndex        =   4
            Top             =   30
            Width           =   3045
            _cx             =   5371
            _cy             =   6720
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   15329975
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   16645618
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmProducts.frx":0106
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
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
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid grdSub1 
            Height          =   3810
            Left            =   2550
            TabIndex        =   5
            Top             =   30
            Width           =   2805
            _cx             =   4948
            _cy             =   6720
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   15329975
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   16645618
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   700
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmProducts.frx":017E
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   5
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
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
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid grdMajor1 
            Height          =   3810
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   3105
            _cx             =   5477
            _cy             =   6720
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   15329975
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   16645618
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   700
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmProducts.frx":01F6
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   5
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
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
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox PicBC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4740
         ScaleHeight     =   255
         ScaleWidth      =   315
         TabIndex        =   89
         Top             =   990
         Visible         =   0   'False
         Width           =   345
         Begin VB.Label lblCode 
            BackColor       =   &H00C0FFC0&
            Caption         =   "BC"
            Height          =   165
            Left            =   60
            TabIndex        =   90
            Top             =   30
            Width           =   285
         End
      End
      Begin VB.TextBox txtUnitSize 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2905
         TabIndex        =   84
         Top             =   2425
         Width           =   2115
      End
      Begin VB.TextBox txtShort 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2905
         MaxLength       =   25
         TabIndex        =   83
         Top             =   1725
         Width           =   2115
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2905
         MaxLength       =   50
         TabIndex        =   82
         Top             =   1375
         Width           =   5535
      End
      Begin VB.TextBox txtProductCode 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2905
         TabIndex        =   69
         Top             =   1030
         Width           =   1785
      End
      Begin btButtonEx.ButtonEx cmdLinks 
         Height          =   315
         Left            =   12045
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         Appearance      =   3
         Caption         =   "View Recipe Links"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtSellIncl 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2860
         TabIndex        =   76
         Text            =   "0.00"
         Top             =   5170
         Width           =   1725
      End
      Begin VB.TextBox txtSellExcl 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2860
         TabIndex        =   75
         Text            =   "0.00"
         Top             =   4490
         Width           =   1725
      End
      Begin VB.TextBox txtGross 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   2860
         TabIndex        =   74
         Text            =   "0"
         Top             =   4130
         Width           =   1725
      End
      Begin VB.TextBox txtMarkup 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2860
         TabIndex        =   73
         Text            =   "0"
         Top             =   3745
         Width           =   1725
      End
      Begin VB.TextBox txtLandCost 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2860
         TabIndex        =   72
         Text            =   "0.00"
         Top             =   3390
         Width           =   1725
      End
      Begin VB.PictureBox picTopFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   0
         ScaleHeight     =   585
         ScaleWidth      =   13785
         TabIndex        =   38
         Top             =   5640
         Width           =   13785
         Begin btButtonEx.ButtonEx ButtonEx4 
            Height          =   345
            Left            =   5880
            TabIndex        =   39
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            Appearance      =   3
            Caption         =   "Refresh"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin btButtonEx.ButtonEx cmdUp 
            Height          =   465
            Left            =   13080
            TabIndex        =   40
            Top             =   60
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   820
            Appearance      =   3
            Caption         =   "5"
            CaptionOffsetX  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Webdings"
               Size            =   15.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocus       =   0
         End
         Begin btButtonEx.ButtonEx cmdSearch 
            Height          =   345
            Left            =   4530
            TabIndex        =   92
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            Appearance      =   3
            Caption         =   "Search"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblUsers 
            Caption         =   "Product List."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   150
            TabIndex        =   43
            Top             =   60
            Width           =   3135
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   2
            Left            =   3780
            Top             =   420
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   1
            Left            =   3450
            Top             =   420
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   0
            Left            =   3120
            Top             =   420
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin MSForms.Image Image3 
            Height          =   90
            Left            =   60
            Top             =   420
            Width           =   3015
            BackColor       =   16761024
            Size            =   "5318;159"
         End
         Begin MSForms.ComboBox cmbDepart 
            Height          =   375
            Left            =   10020
            TabIndex        =   42
            Top             =   90
            Width           =   2925
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "5159;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Arial Narrow"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbType 
            Height          =   375
            Left            =   7290
            TabIndex        =   41
            Top             =   90
            Width           =   2625
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "4630;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Arial Narrow"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Image picTopbar 
            Height          =   675
            Left            =   0
            Top             =   0
            Width           =   13755
            BorderStyle     =   0
            SpecialEffect   =   3
            Size            =   "24262;1191"
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   240
         Top             =   240
      End
      Begin VSFlex8Ctl.VSFlexGrid grdMenu 
         Height          =   3615
         Left            =   30
         TabIndex        =   1
         Top             =   1200
         Visible         =   0   'False
         Width           =   150
         _cx             =   265
         _cy             =   6376
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin btButtonEx.ButtonEx ButtonEx1 
         Height          =   285
         Left            =   1260
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2730
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Appearance      =   3
         Caption         =   "Department..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignHorz=   0
      End
      Begin BTNENHLib4.BtnEnh cmdTab 
         Height          =   420
         Index           =   3
         Left            =   11490
         TabIndex        =   45
         Top             =   1740
         Width           =   2100
         _Version        =   524298
         _ExtentX        =   3695
         _ExtentY        =   741
         _StockProps     =   66
         Caption         =   "Images"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Shape           =   4
         Surface         =   1
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmProducts.frx":026E
         textLT          =   "frmProducts.frx":02DA
         textCT          =   "frmProducts.frx":02F2
         textRT          =   "frmProducts.frx":030A
         textLM          =   "frmProducts.frx":0322
         textRM          =   "frmProducts.frx":033A
         textLB          =   "frmProducts.frx":0352
         textCB          =   "frmProducts.frx":036A
         textRB          =   "frmProducts.frx":0382
         colorBack       =   "frmProducts.frx":039A
         colorIntern     =   "frmProducts.frx":03C4
         colorMO         =   "frmProducts.frx":03EE
         colorFocus      =   "frmProducts.frx":0418
         colorDisabled   =   "frmProducts.frx":0442
         colorPressed    =   "frmProducts.frx":046C
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   5
      End
      Begin BTNENHLib4.BtnEnh cmdTab 
         Height          =   420
         Index           =   2
         Left            =   9405
         TabIndex        =   46
         Top             =   1740
         Width           =   2115
         _Version        =   524298
         _ExtentX        =   3731
         _ExtentY        =   741
         _StockProps     =   66
         Caption         =   "Suppliers"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Shape           =   4
         Surface         =   1
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmProducts.frx":0496
         textLT          =   "frmProducts.frx":0508
         textCT          =   "frmProducts.frx":0520
         textRT          =   "frmProducts.frx":0538
         textLM          =   "frmProducts.frx":0550
         textRM          =   "frmProducts.frx":0568
         textLB          =   "frmProducts.frx":0580
         textCB          =   "frmProducts.frx":0598
         textRB          =   "frmProducts.frx":05B0
         colorBack       =   "frmProducts.frx":05C8
         colorIntern     =   "frmProducts.frx":05F2
         colorMO         =   "frmProducts.frx":061C
         colorFocus      =   "frmProducts.frx":0646
         colorDisabled   =   "frmProducts.frx":0670
         colorPressed    =   "frmProducts.frx":069A
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   5
      End
      Begin BTNENHLib4.BtnEnh cmdTab 
         Height          =   420
         Index           =   1
         Left            =   7350
         TabIndex        =   47
         Top             =   1740
         Width           =   2085
         _Version        =   524298
         _ExtentX        =   3678
         _ExtentY        =   741
         _StockProps     =   66
         Caption         =   "Recipe"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Shape           =   4
         Surface         =   1
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmProducts.frx":06C4
         textLT          =   "frmProducts.frx":0730
         textCT          =   "frmProducts.frx":0748
         textRT          =   "frmProducts.frx":0760
         textLM          =   "frmProducts.frx":0778
         textRM          =   "frmProducts.frx":0790
         textLB          =   "frmProducts.frx":07A8
         textCB          =   "frmProducts.frx":07C0
         textRB          =   "frmProducts.frx":07D8
         colorBack       =   "frmProducts.frx":07F0
         colorIntern     =   "frmProducts.frx":081A
         colorMO         =   "frmProducts.frx":0844
         colorFocus      =   "frmProducts.frx":086E
         colorDisabled   =   "frmProducts.frx":0898
         colorPressed    =   "frmProducts.frx":08C2
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   5
      End
      Begin BTNENHLib4.BtnEnh cmdTab 
         Height          =   420
         Index           =   0
         Left            =   5295
         TabIndex        =   48
         Top             =   1740
         Width           =   2085
         _Version        =   524298
         _ExtentX        =   3678
         _ExtentY        =   741
         _StockProps     =   66
         Caption         =   "Settings"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Shape           =   4
         Surface         =   1
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmProducts.frx":08EC
         textLT          =   "frmProducts.frx":095C
         textCT          =   "frmProducts.frx":0974
         textRT          =   "frmProducts.frx":098C
         textLM          =   "frmProducts.frx":09A4
         textRM          =   "frmProducts.frx":09BC
         textLB          =   "frmProducts.frx":09D4
         textCB          =   "frmProducts.frx":09EC
         textRB          =   "frmProducts.frx":0A04
         colorBack       =   "frmProducts.frx":0A1C
         colorIntern     =   "frmProducts.frx":0A46
         colorMO         =   "frmProducts.frx":0A70
         colorFocus      =   "frmProducts.frx":0A9A
         colorDisabled   =   "frmProducts.frx":0AC4
         colorPressed    =   "frmProducts.frx":0AEE
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   5
         Value           =   -1  'True
      End
      Begin btButtonEx.ButtonEx cmdBarcode 
         Height          =   285
         Left            =   5190
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   990
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         Appearance      =   3
         Caption         =   "Change Barcode..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignHorz=   0
      End
      Begin btButtonEx.ButtonEx ButtonEx2 
         Height          =   285
         Left            =   6750
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   990
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         Appearance      =   3
         Caption         =   "Print Labels..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignHorz=   0
      End
      Begin btButtonEx.ButtonEx ButtonEx3 
         Height          =   285
         Left            =   900
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   5130
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         Appearance      =   3
         Caption         =   "Selling Prices (incl)..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignHorz=   0
      End
      Begin btButtonEx.ButtonEx cmdTestR 
         Height          =   315
         Left            =   12050
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         Appearance      =   3
         Caption         =   "Test Recipe"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdPrep 
         Height          =   315
         Left            =   12040
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         Appearance      =   3
         Caption         =   "Preparation..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox picTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3375
         Index           =   0
         Left            =   5310
         ScaleHeight     =   3375
         ScaleWidth      =   8295
         TabIndex        =   8
         Top             =   2130
         Width           =   8295
         Begin VB.PictureBox picBox 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3210
            Index           =   0
            Left            =   30
            ScaleHeight     =   3210
            ScaleWidth      =   8205
            TabIndex        =   9
            Top             =   30
            Width           =   8205
            Begin VB.Frame fmLink 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Unit Link"
               ForeColor       =   &H00000000&
               Height          =   1125
               Left            =   6270
               TabIndex        =   93
               Top             =   1980
               Width           =   1845
               Begin VB.TextBox txtLink 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   225
                  Left            =   270
                  Locked          =   -1  'True
                  TabIndex        =   95
                  Text            =   "<Not Linked>"
                  Top             =   315
                  Width           =   1395
               End
               Begin btButtonEx.ButtonEx cmdLink 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   94
                  TabStop         =   0   'False
                  Top             =   630
                  Width           =   1605
                  _ExtentX        =   2831
                  _ExtentY        =   556
                  Appearance      =   3
                  Caption         =   "Link Code..."
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ShowFocus       =   0
               End
               Begin MSForms.Image picLink 
                  Height          =   285
                  Left            =   120
                  Top             =   270
                  Width           =   1605
                  BackColor       =   16777215
                  Size            =   "2831;503"
               End
            End
            Begin VB.CheckBox chkProduction 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Production Item for In-House Manufacturing "
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   91
               Top             =   2880
               Width           =   3615
            End
            Begin VB.TextBox txtPackSize 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   4980
               TabIndex        =   87
               Text            =   "1"
               Top             =   2475
               Width           =   1155
            End
            Begin VB.TextBox txtRef 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   4980
               TabIndex        =   22
               Top             =   1005
               Width           =   3105
            End
            Begin VB.TextBox txtEmpty 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   4980
               TabIndex        =   25
               Top             =   2115
               Width           =   1155
            End
            Begin VB.TextBox txtFull 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   4980
               TabIndex        =   24
               Top             =   1755
               Width           =   1155
            End
            Begin btButtonEx.ButtonEx cmdSuppliers 
               Height          =   315
               Left            =   4830
               TabIndex        =   80
               Top             =   2790
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Appearance      =   3
               Caption         =   "Supplier Codes..."
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.CheckBox chkWhole 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Manage Stock as Whole Units Only"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   18
               Top             =   2535
               Width           =   3135
            End
            Begin VB.CheckBox chkScale 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Scale Item"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   1830
               Width           =   2385
            End
            Begin VB.CheckBox chkDelete 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Delete when Stock runs out"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   2190
               Width           =   2385
            End
            Begin VB.CheckBox chkTouch 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Display on Touch"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   1140
               Width           =   1875
            End
            Begin VB.CheckBox chkRecipe 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Recipe Item"
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   795
               Width           =   1875
            End
            Begin VB.CheckBox chkSales 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Sales Item"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   11
               Top             =   435
               Value           =   1  'Checked
               Width           =   1875
            End
            Begin VB.CheckBox chkStock 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Stock Keeping Item"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   90
               Value           =   1  'Checked
               Width           =   1875
            End
            Begin VB.CheckBox chkDeposit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Deposit Bearing Item"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   1485
               Width           =   2385
            End
            Begin MSForms.ComboBox cmbembedtype 
               Height          =   315
               Left            =   6660
               TabIndex        =   98
               Tag             =   "Up"
               Top             =   1320
               Width           =   1455
               VariousPropertyBits=   746604569
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "2566;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               BorderColor     =   8421504
               SpecialEffect   =   0
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label lblembedtype 
               Height          =   255
               Left            =   6240
               TabIndex        =   97
               Top             =   1350
               Width           =   375
               BackColor       =   -2147483643
               VariousPropertyBits=   8388633
               Caption         =   "Type:"
               Size            =   "661;450"
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Line Line3 
               X1              =   6180
               X2              =   6270
               Y1              =   2610
               Y2              =   2610
            End
            Begin VB.Line Line2 
               X1              =   6180
               X2              =   6270
               Y1              =   2520
               Y2              =   2520
            End
            Begin MSForms.Label Label2 
               Height          =   225
               Index           =   17
               Left            =   3210
               TabIndex        =   88
               Top             =   2475
               Width           =   1545
               BackColor       =   -2147483643
               Caption         =   "Pack Size:"
               Size            =   "2725;397"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.Image Image12 
               Height          =   285
               Left            =   4830
               Top             =   2430
               Width           =   1365
               BackColor       =   16777215
               Size            =   "2408;503"
            End
            Begin MSForms.Label Label2 
               Height          =   225
               Index           =   16
               Left            =   3210
               TabIndex        =   85
               Top             =   1005
               Width           =   1545
               BackColor       =   -2147483643
               Caption         =   "Nappi Code:"
               Size            =   "2725;397"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.Image Image10 
               Height          =   285
               Index           =   1
               Left            =   4830
               Top             =   960
               Width           =   3285
               BackColor       =   16777215
               Size            =   "5794;503"
            End
            Begin MSForms.Image Image11 
               Height          =   285
               Left            =   4830
               Top             =   2070
               Width           =   1365
               BackColor       =   16777215
               Size            =   "2408;503"
            End
            Begin MSForms.Image Image10 
               Height          =   285
               Index           =   0
               Left            =   4830
               Top             =   1710
               Width           =   1365
               BackColor       =   16777215
               Size            =   "2408;503"
            End
            Begin MSForms.Label Label2 
               Height          =   225
               Index           =   14
               Left            =   3210
               TabIndex        =   79
               Top             =   2115
               Width           =   1545
               BackColor       =   -2147483643
               Caption         =   "Empty Bottle Weight:"
               Size            =   "2725;397"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.Label Label2 
               Height          =   225
               Index           =   1
               Left            =   3210
               TabIndex        =   78
               Top             =   1755
               Width           =   1545
               BackColor       =   -2147483643
               Caption         =   "Full Bottle Weight:"
               Size            =   "2725;397"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.Label lblScale 
               Height          =   255
               Left            =   3150
               TabIndex        =   26
               Top             =   1380
               Width           =   1605
               BackColor       =   -2147483643
               VariousPropertyBits=   8388633
               Caption         =   "Scale Prefix:"
               Size            =   "2831;450"
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.ComboBox cmbScalePrefix 
               Height          =   315
               Left            =   4830
               TabIndex        =   23
               Tag             =   "Up"
               Top             =   1320
               Width           =   1365
               VariousPropertyBits=   746604569
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "2408;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               BorderColor     =   8421504
               SpecialEffect   =   0
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbPrinter2 
               Height          =   315
               Left            =   4830
               TabIndex        =   21
               Tag             =   "Up"
               Top             =   585
               Width           =   3285
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5794;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label2 
               Height          =   255
               Index           =   15
               Left            =   3150
               TabIndex        =   20
               Top             =   615
               Width           =   1605
               BackColor       =   -2147483643
               Caption         =   "Kitchen Printer no.2:"
               Size            =   "2831;450"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.ComboBox cmbPrinter1 
               Height          =   315
               Left            =   4830
               TabIndex        =   19
               Tag             =   "Up"
               Top             =   180
               Width           =   3285
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5794;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label2 
               Height          =   255
               Index           =   8
               Left            =   3150
               TabIndex        =   16
               Top             =   210
               Width           =   1605
               BackColor       =   -2147483643
               Caption         =   "Kitchen Printer no.1:"
               Size            =   "2831;450"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
         End
         Begin VB.PictureBox picBox 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3195
            Index           =   1
            Left            =   60
            ScaleHeight     =   3195
            ScaleWidth      =   8175
            TabIndex        =   36
            Top             =   60
            Width           =   8175
            Begin VSFlex8Ctl.VSFlexGrid grdRecipe 
               Height          =   3150
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Width           =   8160
               _cx             =   14393
               _cy             =   5556
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   15329975
               ForeColorSel    =   0
               BackColorBkg    =   16777215
               BackColorAlternate=   16645618
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483632
               FocusRect       =   0
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmProducts.frx":0B18
               ScrollTrack     =   -1  'True
               ScrollBars      =   2
               ScrollTips      =   0   'False
               MergeCells      =   2
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
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
               ComboSearch     =   1
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
         End
         Begin VB.PictureBox picBox 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3045
            Index           =   3
            Left            =   60
            ScaleHeight     =   3045
            ScaleWidth      =   8190
            TabIndex        =   29
            Top             =   60
            Width           =   8195
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Touch Panel Image"
               ForeColor       =   &H80000008&
               Height          =   3045
               Left            =   4290
               TabIndex        =   30
               Top             =   -30
               Width           =   3705
               Begin BTNENHLib4.BtnEnh BtnEnh1 
                  Height          =   1125
                  Index           =   1
                  Left            =   180
                  TabIndex        =   31
                  Top             =   1620
                  Width           =   1935
                  _Version        =   524298
                  _ExtentX        =   3413
                  _ExtentY        =   1984
                  _StockProps     =   66
                  Caption         =   "Test"
                  BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Shape           =   1
                  CornerFactor    =   10
                  Surface         =   1
                  PicturePosition =   10
                  SmoothEdges     =   1
                  SpecialEffect   =   1
                  LogPixels       =   96
                  SpecialEffectFactor=   3
                  TextureBevelFactor=   4
                  UserData        =   0.1
                  textCaption     =   "frmProducts.frx":0B90
                  textLT          =   "frmProducts.frx":0BF8
                  textCT          =   "frmProducts.frx":0C10
                  textRT          =   "frmProducts.frx":0C28
                  textLM          =   "frmProducts.frx":0C40
                  textRM          =   "frmProducts.frx":0C58
                  textLB          =   "frmProducts.frx":0C70
                  textCB          =   "frmProducts.frx":0C88
                  textRB          =   "frmProducts.frx":0CA0
                  colorBack       =   "frmProducts.frx":0CB8
                  colorIntern     =   "frmProducts.frx":0CE2
                  colorMO         =   "frmProducts.frx":0D0C
                  colorFocus      =   "frmProducts.frx":0D36
                  colorDisabled   =   "frmProducts.frx":0D60
                  colorPressed    =   "frmProducts.frx":0D8A
                  Style           =   1
                  HollowFrame     =   -1  'True
                  LightDirection  =   1
                  Value           =   -1  'True
               End
               Begin BTNENHLib4.BtnEnh BtnEnh1 
                  Height          =   1125
                  Index           =   0
                  Left            =   180
                  TabIndex        =   32
                  Top             =   390
                  Width           =   1935
                  _Version        =   524298
                  _ExtentX        =   3413
                  _ExtentY        =   1984
                  _StockProps     =   66
                  Caption         =   "Test"
                  BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Shape           =   1
                  CornerFactor    =   10
                  Surface         =   1
                  PicturePosition =   10
                  SmoothEdges     =   1
                  SpecialEffect   =   1
                  LogPixels       =   96
                  SpecialEffectFactor=   3
                  TextureBevelFactor=   4
                  UserData        =   0.1
                  textCaption     =   "frmProducts.frx":0DB4
                  textLT          =   "frmProducts.frx":0E1C
                  textCT          =   "frmProducts.frx":0E34
                  textRT          =   "frmProducts.frx":0E4C
                  textLM          =   "frmProducts.frx":0E64
                  textRM          =   "frmProducts.frx":0E7C
                  textLB          =   "frmProducts.frx":0E94
                  textCB          =   "frmProducts.frx":0EAC
                  textRB          =   "frmProducts.frx":0EC4
                  colorBack       =   "frmProducts.frx":0EDC
                  colorIntern     =   "frmProducts.frx":0F06
                  colorMO         =   "frmProducts.frx":0F30
                  colorFocus      =   "frmProducts.frx":0F5A
                  colorDisabled   =   "frmProducts.frx":0F84
                  colorPressed    =   "frmProducts.frx":0FAE
                  HollowFrame     =   -1  'True
                  LightDirection  =   1
               End
               Begin MSForms.Label Label9 
                  Height          =   795
                  Left            =   2670
                  TabIndex        =   34
                  Top             =   1770
                  Width           =   765
                  ForeColor       =   8421504
                  BackColor       =   -2147483643
                  Caption         =   "Down State"
                  Size            =   "1349;1402"
                  FontName        =   "Arial Narrow"
                  FontEffects     =   1073741825
                  FontHeight      =   285
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   3
                  FontWeight      =   700
               End
               Begin MSForms.Label Label8 
                  Height          =   645
                  Left            =   2670
                  TabIndex        =   33
                  Top             =   720
                  Width           =   765
                  ForeColor       =   8421504
                  BackColor       =   -2147483643
                  Caption         =   "Up State"
                  Size            =   "1349;1138"
                  FontName        =   "Arial Narrow"
                  FontEffects     =   1073741825
                  FontHeight      =   285
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   3
                  FontWeight      =   700
               End
            End
            Begin MSForms.Image Image13 
               Height          =   2955
               Left            =   90
               Top             =   60
               Width           =   4035
               BackColor       =   16777215
               Size            =   "7117;5212"
            End
            Begin MSForms.Label Label1 
               Height          =   885
               Left            =   960
               TabIndex        =   35
               Top             =   990
               Width           =   2175
               ForeColor       =   14737632
               BackColor       =   -2147483643
               Caption         =   "Click to add Picture"
               Size            =   "3836;1561"
               FontName        =   "Arial"
               FontEffects     =   1073741825
               FontHeight      =   360
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
         End
         Begin VB.PictureBox picBox 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3250
            Index           =   2
            Left            =   0
            ScaleHeight     =   3255
            ScaleWidth      =   8250
            TabIndex        =   27
            Top             =   30
            Width           =   8250
            Begin VSFlex8Ctl.VSFlexGrid grdSuppliers 
               Height          =   3260
               Left            =   20
               TabIndex        =   28
               Top             =   0
               Width           =   8240
               _cx             =   14534
               _cy             =   5733
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   15329975
               ForeColorSel    =   0
               BackColorBkg    =   16777215
               BackColorAlternate=   16645618
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483632
               FocusRect       =   0
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmProducts.frx":0FD8
               ScrollTrack     =   -1  'True
               ScrollBars      =   2
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   1
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   5
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
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
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
         End
         Begin MSForms.Image Image2 
            Height          =   3330
            Left            =   0
            Top             =   -30
            Width           =   8265
            BorderColor     =   8421504
            BackColor       =   16777215
            Size            =   "14579;5874"
         End
      End
      Begin btButtonEx.ButtonEx cmdPrint 
         Height          =   315
         Left            =   12030
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         Appearance      =   3
         Caption         =   "Print Recipe"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Txtstocklevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   10260
         TabIndex        =   100
         Top             =   960
         Width           =   1455
      End
      Begin MSForms.Image Image15 
         Height          =   315
         Left            =   10020
         Top             =   930
         Width           =   1935
         BackColor       =   16777215
         Size            =   "3413;556"
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   19
         Left            =   8790
         TabIndex        =   99
         Top             =   960
         Width           =   1215
         BackColor       =   -2147483643
         Caption         =   "Stock Level min:"
         Size            =   "2143;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Image Image7 
         Height          =   285
         Index           =   1
         Left            =   2760
         Top             =   2380
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;503"
      End
      Begin MSForms.Image Image9 
         Height          =   285
         Left            =   2760
         Top             =   1680
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;503"
      End
      Begin MSForms.Image Image8 
         Height          =   285
         Left            =   2760
         Top             =   1330
         Width           =   5745
         BackColor       =   16777215
         Size            =   "10134;503"
      End
      Begin MSForms.Image Image7 
         Height          =   285
         Index           =   0
         Left            =   2760
         Top             =   990
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;503"
      End
      Begin MSForms.Image picSeperate 
         Height          =   135
         Left            =   -60
         Top             =   5550
         Width           =   13815
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "24368;238"
      End
      Begin VB.Label lblOrderFilter 
         BackColor       =   &H00FFFFFF&
         Caption         =   "order by Description"
         Height          =   525
         Left            =   5790
         TabIndex        =   77
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   4
         Left            =   2760
         Top             =   5110
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   3
         Left            =   2760
         Top             =   3330
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   2
         Left            =   2760
         Top             =   4410
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   1
         Left            =   2760
         Top             =   4050
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image6 
         Height          =   300
         Left            =   2760
         Top             =   3330
         Width           =   2325
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image1 
         Height          =   345
         Left            =   840
         Top             =   210
         Width           =   3195
         BackColor       =   16777215
         Size            =   "5636;609"
         VariousPropertyBits=   19
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   900
         TabIndex        =   71
         Top             =   240
         Width           =   3105
         ForeColor       =   0
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Product Details"
         Size            =   "5477;450"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   0
         Left            =   1260
         TabIndex        =   70
         Top             =   2400
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Unit Size:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblCost 
         Height          =   225
         Left            =   1260
         TabIndex        =   68
         Top             =   3420
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Landed Cost:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   1260
         TabIndex        =   67
         Top             =   3765
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Markup%"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   3
         Left            =   1260
         TabIndex        =   66
         Top             =   4110
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Gross Profit%:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   4
         Left            =   1260
         TabIndex        =   65
         Top             =   4455
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Selling Price (excl):"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   5
         Left            =   1260
         TabIndex        =   64
         Top             =   4800
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Tax:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   6
         Left            =   1260
         TabIndex        =   63
         Top             =   1050
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Product Code:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.ComboBox cmbTax 
         Height          =   285
         Left            =   2760
         TabIndex        =   62
         Tag             =   "Up"
         Top             =   4770
         Width           =   2325
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4101;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbDepartments 
         Height          =   285
         Left            =   2760
         TabIndex        =   61
         Tag             =   "Up"
         Top             =   2730
         Width           =   2325
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4101;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbUnit 
         Height          =   285
         Left            =   2760
         TabIndex        =   60
         Tag             =   "Up"
         Top             =   2040
         Width           =   2325
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4101;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   59
         Top             =   3180
         Width           =   1185
         ForeColor       =   12582912
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Pricing"
         Size            =   "2090;450"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1320
         X2              =   5100
         Y1              =   3210
         Y2              =   3210
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   90
         TabIndex        =   58
         Top             =   690
         Width           =   1185
         ForeColor       =   12582912
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "General"
         Size            =   "2090;450"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1320
         X2              =   8580
         Y1              =   750
         Y2              =   750
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   10
         Left            =   1260
         TabIndex        =   57
         Top             =   2085
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Unit of Measure:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   11
         Left            =   1260
         TabIndex        =   56
         Top             =   1740
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Button Description:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   12
         Left            =   1260
         TabIndex        =   55
         Top             =   1395
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Description:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblRange 
         Height          =   255
         Left            =   10110
         TabIndex        =   54
         Top             =   1380
         Width           =   1875
         BackColor       =   -2147483643
         VariousPropertyBits=   8388627
         Caption         =   "0 to 0"
         Size            =   "3307;450"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   18
         Left            =   8790
         TabIndex        =   53
         Top             =   1365
         Width           =   1215
         BackColor       =   -2147483643
         Caption         =   "Product Range:"
         Size            =   "2143;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   0
         Left            =   2760
         Top             =   3690
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image14 
         Height          =   315
         Left            =   10050
         Top             =   1320
         Width           =   1935
         BackColor       =   16777215
         Size            =   "3413;556"
      End
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkTouch_Click()
' inproducts
If chkTouch.Value = 1 Then
cmbScalePrefix.Text = "<None>"
cmbembedtype.Text = "<None>"
End If
   
End Sub

Private Sub cmbTax_Change()
    Calculate_Plu "Tax"
End Sub

Private Sub cmdLink_Click()
    Load frmSearch
    frmSearch.Tag = "Links"
    frmSearch.Show vbModal
    If frmSearch.Tag = "" Then
        txtLink.Text = "<Not Linked>"
    Else
        txtLink.Text = Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1))
        frmSearch.Tag = ""
    End If
    Unload frmSearch
End Sub
Private Sub cmdLinks_Click()
    frmRecipeLinks.Show vbModal
End Sub
Private Sub cmdPrep_Click()
    frmPrepMethod.Show vbModal
End Sub

Private Sub cmdPrint_Click()
    ActiveUpdateServer "Delete from Rec_Temp"
    DoEvents
    With frmProducts
        For i = 1 To frmProducts.grdRecipe.Rows - 1
            ActiveUpdateServer "INSERT INTO [Rec_Temp]([PType], [Description], [Unit], [Qty],[Cost])" & _
            " VALUES('" & .grdRecipe.TextMatrix(i, 0) & "','" & .grdRecipe.TextMatrix(i, 1) & "','" & .grdRecipe.TextMatrix(i, 2) & "','" & .grdRecipe.TextMatrix(i, 3) & "'," & .grdRecipe.TextMatrix(i, 4) & ")"
        Next i
    End With
    frmProducts.Tag = "Not Now"
    DoEvents
    rptRecipe.Show vbModal
End Sub
Private Sub cmdSearch_Click()
    ProductFilter(2) = ""
    frmProdFind.Show vbModal
End Sub
Private Sub cmdSuppliers_Click()
    frmSuppCodes.Show vbModal
End Sub
Private Sub Form_Activate()
    If frmProducts.Tag = "Not Now" Then
        frmProducts.Tag = ""
        Exit Sub
    End If
    Screen.MousePointer = 11
    On Error Resume Next
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Products"
    frmMain.Toolbar1.Buttons(2).Caption = "New"
    frmMain.picProdBar.Visible = True
    txtProductCode.SetFocus
    If ReplicationServ = 1 Then
        txtProductCode.Locked = True
        txtDescription.Locked = True
        txtShort.Locked = True
        cmbUnit.Locked = True
        txtUnitSize.Locked = True
        cmbDepartments.Locked = True
        txtLandCost.Locked = True
        txtMarkup.Locked = True
        txtGross.Locked = True
        txtSellExcl.Locked = True
        cmbTax.Locked = True
        txtSellExcl.Locked = True
        txtSellIncl.Locked = True
        cmdBarcode.Enabled = False
        ButtonEx3.Enabled = False
        picTab(0).Enabled = False
        grdPrices.Enabled = False
        txtPackSize.Enabled = False
    End If
    txtProductCode.SelStart = Len(txtProductCode.Text)
    grdProd.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            KeyCode = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
            frmSplash.Show
            frmMain.picProdBar.Visible = False
            frmMain.picAccBar.Visible = False
            frmMain.Hide
    End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
End Sub
Private Sub grdProd_BeforeSort(ByVal Col As Long, Order As Integer)
    On Error Resume Next
    Select Case Trim(grdProd.TextMatrix(0, Col))
        Case "Product Code"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by Product_Code"
                Case 2
                    lblOrderFilter.Caption = " Order by Product_Code Desc"
            End Select
        Case "Description"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [Description]"
                Case 2
                    lblOrderFilter.Caption = " Order by [Description] Desc"
            End Select
        Case "Department"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [Dept_Name]"
                Case 2
                    lblOrderFilter.Caption = " Order by [Dept_Name] Desc"
            End Select
        Case "Stock on Hand"
             Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [SOH]"
                Case 2
                    lblOrderFilter.Caption = " Order by [SOH] Desc"
            End Select
        Case "Landed Cost"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [Landed_Cost]"
                Case 2
                    lblOrderFilter.Caption = " Order by [Landed_Cost] Desc"
            End Select
         Case "Tax Rate"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [Tax_Rate]"
                Case 2
                    lblOrderFilter.Caption = " Order by [Tax_Rate] Desc"
            End Select
        Case "Price (Incl)"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [Selling_Price]"
                Case 2
                    lblOrderFilter.Caption = " Order by [Selling_Price] Desc"
            End Select
    End Select
    On Error GoTo 0
End Sub

Private Sub grdSuppliers_DblClick()
    TillData.DocNo = grdSuppliers.TextMatrix(grdSuppliers.Row, 5)
    frmProducts.Tag = "Not Now"
    rptGRV1.Show vbModal
    TillData.DocNo = 0
End Sub



Private Sub txtEmpty_GotFocus()
    txtEmpty.SelStart = 0
    txtEmpty.SelLength = Len(txtEmpty.Text)
End Sub

Private Sub txtEmpty_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtFull_GotFocus()
    txtFull.SelStart = 0
    txtFull.SelLength = Len(txtFull.Text)
End Sub

Private Sub txtFull_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtGross_Change()
    If txtGross.Tag = "" Then Calculate_Plu "GP"
End Sub
Private Sub txtLandCost_Change()
    Calculate_Plu "Landed Cost"
End Sub
Private Sub txtMarkup_Change()
    If txtMarkup.Tag = "" Then Calculate_Plu "Markup"
End Sub
Private Sub txtPackSize_GotFocus()
    txtPackSize.SelStart = 0
    txtPackSize.SelLength = Len(txtPackSize.Text)
End Sub
Private Sub txtPackSize_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtRef_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub
Private Sub txtRef_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            cmbPrinter2.SetFocus
        Case 40
            txtShort.SetFocus
    End Select
End Sub
Private Sub txtRef_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub
Private Sub txtRef_LostFocus()
    On Error Resume Next
    CheckforSave
    On Error GoTo 0
End Sub
Private Sub txtSellExcl_Change()
    If txtSellExcl.Tag = "" Then Calculate_Plu "SellExcl"
End Sub

Private Sub txtSellIncl_Change()
    If txtSellIncl.Tag = "" Then Calculate_Plu "SellIncl"
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub


Private Sub txtShort_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub

Private Sub txtShort_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtDescription.SetFocus
        Case 40
            cmbUnit.SetFocus
    End Select
End Sub
Private Sub txtShort_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 33 To 47, 58 To 64, 91 To 96, 123 To 127, 162 To 184, 247, 248, 191
            KeyAscii = 0
        
    End Select
End Sub

Private Sub Txtstocklevel_Change()
If chkStock.Value = 0 Then Txtstocklevel.Text = "None"
End Sub

Private Sub txtUnitSize_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub

Private Sub txtUnitSize_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            cmbUnit.SetFocus
        Case 40
            cmbDepartments.SetFocus
    End Select
End Sub

Private Sub txtUnitSize_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub cmdTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    For i = 0 To cmdTab.Count - 1
        If Index = i Then
            picBox(i).Visible = True
        Else
            picBox(i).Visible = False
        End If
        cmdTestR.Visible = False
        cmdPrep.Visible = True
    Next i
    If Index = 1 Then
        cmdTestR.Visible = True
        cmdPrep.Visible = False
    End If
    If Index = 1 Then cmdLinks.Visible = False
    If Index = 0 Then
        If Abs(chkStock.Value) = 1 Then
            cmdLinks.Visible = True
        End If
    End If
    If Index = 2 Then
        LoadSuppInfo
    End If
End Sub
Private Sub LoadSuppInfo()
    Screen.MousePointer = 11
    ActiveReadServer3 "SELECT Purchase_Journal.Grv_No,Purchase_Journal.Supplier_No, Suppliers.Supplier_Name, Purchase_Journal.Qty_Invoiced," & _
    " Purchase_Journal.Invoice_Date , Purchase_Journal.Price_Invoiced, Purchase_Journal.Product_Code" & _
    " FROM Purchase_Journal LEFT OUTER JOIN" & _
    " Suppliers ON Purchase_Journal.Supplier_No = Suppliers.Supplier_No" & _
    " WHERE (Purchase_Journal.Product_Code = '" & txtProductCode.Text & "')" & _
    " ORDER BY Purchase_Journal.Invoice_Date DESC"
    grdSuppliers.Rows = 1
    grdSuppliers.ColHidden(5) = True
    While Not rs3.EOF
        grdSuppliers.Rows = grdSuppliers.Rows + 1
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 0) = rs3.Fields("Supplier_No")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 1) = rs3.Fields("Supplier_Name")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 2) = rs3.Fields("Qty_Invoiced")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 3) = Format(rs3.Fields("Invoice_Date"), "YYYY-MM-DD HH:MM")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 4) = Format(rs3.Fields("Price_Invoiced"), "0.00")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 5) = rs3.Fields("GRV_No")
        rs3.MoveNext
    Wend
    rs3.Close
    Screen.MousePointer = 0
End Sub
Private Sub cmdTestR_Click()
    ActiveUpdateServer "Delete from Recipes where Product_Code='" & txtProductCode.Text & "'"
    DoEvents
    For i = 1 To grdRecipe.Rows - 1
        If grdRecipe.TextMatrix(i, 1) <> "" Then
            Select Case grdRecipe.TextMatrix(i, 0)
                Case "Message"
                    LineType = 0
                Case "Preparation Recipe"
                    LineType = 1
                Case "Sales Item"
                    LineType = 2
                Case "Stock Item"
                    LineType = 3
                Case "Stock Item (Hidden)"
                    LineType = 4
                Case "Price/Size Change"
                    LineType = 5
                Case "Sales Item (Choice)"
                    LineType = 6
                Case "Stock Item (Choice)"
                    LineType = 7
                Case "Exit"
                    LineType = 8
            End Select
            ActiveUpdateServer "INSERT INTO Recipes ([Line_Type], [Product_Code], [Line_Code], [Description], [Unit_of_Measure], [Qty_Used], [Cost])" & _
            "VALUES (" & LineType & "," & txtProductCode.Text & "," & Trim(Val(Trim(Mid(grdRecipe.TextMatrix(i, 1), InStrRev(grdRecipe.TextMatrix(i, 1), ",") + 1)))) & ",'" & grdRecipe.TextMatrix(i, 1) & "','" & Trim(grdRecipe.TextMatrix(i, 2)) & "','" & Trim(grdRecipe.TextMatrix(i, 3)) & "'," & Val(grdRecipe.TextMatrix(i, 4)) & ")"
        Else
            If grdRecipe.TextMatrix(i, 0) = "Exit" Then
                grdRecipe.TextMatrix(i, 1) = "Exit"
                LineType = 8
                ActiveUpdateServer "INSERT INTO Recipes ([Line_Type], [Product_Code], [Line_Code], [Description], [Unit_of_Measure], [Qty_Used], [Cost])" & _
                "VALUES (" & LineType & "," & txtProductCode.Text & "," & Trim(Val(Trim(Mid(grdRecipe.TextMatrix(i, 1), InStrRev(grdRecipe.TextMatrix(i, 1), ",") + 1)))) & ",'" & grdRecipe.TextMatrix(i, 1) & "','" & Trim(grdRecipe.TextMatrix(i, 2)) & "','" & Trim(grdRecipe.TextMatrix(i, 3)) & "'," & Val(grdRecipe.TextMatrix(i, 4)) & ")"
            End If
        End If
    Next i
    grdMenu.Rows = 0
    ActiveReadServer "Select * from recipes where Product_Code= '" & txtProductCode.Text & "' order by Line_No"
    While Not rs.EOF
        grdMenu.Rows = grdMenu.Rows + 1
        grdMenu.TextMatrix(grdMenu.Rows - 1, 0) = rs.Fields("line_no")
        grdMenu.TextMatrix(grdMenu.Rows - 1, 1) = rs.Fields("line_code")
        grdMenu.TextMatrix(grdMenu.Rows - 1, 2) = rs.Fields("line_type")
        grdMenu.TextMatrix(grdMenu.Rows - 1, 3) = Replace(rs.Fields("Description"), "&", "&&")
        rs.MoveNext
    Wend
    rs.Close
    If grdMenu.Rows > 0 Then grdMenu.Row = 0
    Load frmRecipe
    frmRecipe.top = 1000
    frmRecipe.Left = 1000
    frmRecipe.Tag = "Test"
    DoEvents
    frmRecipe.Show vbModal
    frmRecipe.Tag = ""
End Sub

Private Sub cmdUp_Click()
    Select Case cmdUp.Caption
        Case "5"
            grdProd.SetFocus
            DoEvents
            picSeperate.top = 0
            picTopFrame.top = picSeperate.top + picSeperate.Height - 40
            grdProd.top = picSeperate.top + picSeperate.Height - 50 + picTopFrame.Height - 20
            grdProd.Height = picProducts(0).Height - picSeperate.Height - picTopFrame.Height - 550
            cmdUp.Caption = 6
        Case "6"
            picTopFrame.top = 5630
            picSeperate.top = 5550
            grdProd.Height = 2670
            grdProd.top = 6210
            cmdUp.Caption = "5"
            txtProductCode.SetFocus
    End Select
End Sub

Private Sub cmbDepartments_GotFocus()
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub

Private Sub ButtonEx1_Click()
    On Error Resume Next
    If picPrice.Visible = True Then picPrice.Visible = False
    DoEvents
    Select Case picDepartments.Visible
        Case False
            picDepartments.Visible = True
            grdMajor1.TextMatrix(0, 0) = "No."
            grdMajor1.TextMatrix(0, 1) = "Major Department"
            grdSub1.TextMatrix(0, 0) = "No."
            grdSub1.TextMatrix(0, 1) = "Sub Department"
            grdMinor1.TextMatrix(0, 0) = "No."
            grdMinor1.TextMatrix(0, 1) = "Location Link"
            grdMinor1.TextMatrix(0, 2) = "To"
            grdMajor1.ColAlignment(0) = flexAlignLeftCenter
            grdMajor1.ColAlignment(1) = flexAlignLeftCenter
            grdSub1.ColAlignment(0) = flexAlignLeftCenter
            grdSub1.ColAlignment(1) = flexAlignLeftCenter
            grdMinor1.ColAlignment(0) = flexAlignLeftCenter
            grdMinor1.ColAlignment(1) = flexAlignLeftCenter
            grdMinor1.ColAlignment(2) = flexAlignCenterCenter
            grdMinor1.ColDataType(2) = flexDTBoolean
            grdMajor1.Rows = 1
            grdSub1.Rows = 1
            grdMinor1.Rows = 1
            ActiveReadServer "Select * from Locations order by Location_No"
            While Not rs.EOF
                grdMinor1.Rows = grdMinor1.Rows + 1
                grdMinor1.Row = grdMinor1.Rows - 1
                grdMinor1.TextMatrix(grdMinor1.Row, 0) = rs.Fields("Location_No")
                grdMinor1.TextMatrix(grdMinor1.Row, 1) = rs.Fields("Loc_name")
                rs.MoveNext
            Wend
            rs.Close
            If grdMinor1.Rows > 0 Then grdMinor1.Row = 1
            ActiveReadServer "Select * From Departments where Dept_Type=0 order by Department_No"
            i = 0
            While Not rs.EOF
                grdMajor1.Rows = grdMajor1.Rows + 1
                i = i + 1
                grdMajor1.TextMatrix(i, 0) = rs.Fields("Department_No")
                grdMajor1.TextMatrix(i, 1) = rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            If grdMajor1.Rows > 1 Then
                grdSub1.Rows = 1
                ActiveReadServer "Select * From Departments where Dept_Type=1 and Dept_Parent= '" & grdMajor1.TextMatrix(1, 0) & "' order by Department_No"
                i = 0
                While Not rs.EOF
                    grdSub1.Rows = grdSub1.Rows + 1
                    i = i + 1
                    grdSub1.TextMatrix(i, 0) = rs.Fields("Department_No")
                    grdSub1.TextMatrix(i, 1) = rs.Fields("Dept_Name")
                    rs.MoveNext
                Wend
                rs.Close
            End If
            If cmbDepartments.Text <> "<Unbound>" Then
                grdMajor1.Tag = "Not"
                If grdMajor1.FindRow(Trim(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, "-") - 1)), 0, 0, 1, 1) <> -1 Then
                    grdMajor1.Row = grdMajor1.FindRow(Trim(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, "-") - 1)), 0, 0, 1, 1)
                End If
                If grdSub1.FindRow(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, " -") - 1), 0, 0, 1, 1) <> -1 Then
                    grdSub1.Row = grdSub1.FindRow(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, " -") - 1), 0, 0, 1, 1)
                End If
                grdMajor1.Tag = ""
            End If
        Case True
            picDepartments.Visible = False
    End Select
    On Error GoTo 0
End Sub

Public Sub Calculate_Plu(FromWhere)
    On Error Resume Next
    If txtSellIncl.Enabled = False Then Exit Sub
    Select Case FromWhere
        Case "Landed Cost"
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                txtMarkup.Tag = ""
            Else
                 txtMarkup.Text = "0"
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
                txtGross.Tag = ""
            End If
        Case "Markup"
            If Val(txtLandCost.Text) <> 0 Then
                txtSellExcl.Tag = "1"
                txtSellExcl.Text = Format(Val(txtLandCost.Text) * ((100 + Val(txtMarkup.Text)) / 100), "0.00")
                txtSellExcl.Tag = ""
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl) * 100), 3)
                txtGross.Tag = ""
            Else
                txtGross.Tag = "1"
                txtGross.Text = "0"
                txtGross.Tag = ""
            End If
            txtSellIncl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If Val(txtLandCost.Text) <> 0 Then
                txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
            End If
            txtSellIncl.Tag = ""
        Case "GP"
            txtSellExcl.Tag = "1"
            txtSellExcl.Text = Format(txtLandCost.Text / ((100 - Val(txtGross.Text)) / 100), "0.00")
            txtSellExcl.Tag = ""
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                txtMarkup.Tag = ""
            End If
            txtSellIncl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
            txtSellIncl.Tag = ""
        Case "SellExcl"
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost) * 100), 3)
                txtMarkup.Tag = ""
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl) * 100), 3)
                txtGross.Tag = ""
            Else
                txtGross.Tag = "1"
                txtGross.Text = "0"
                txtGross.Tag = ""
            End If
            txtSellIncl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If txtSellExcl.Text <> "N/A" Then
                txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
            End If
            txtSellIncl.Tag = ""
        Case "Tax"
            txtSellIncl.Tag = "1"
            If cmbTax.Text <> "" Then
                Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            End If
            txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
            txtSellIncl.Tag = ""
        Case "SellIncl"
            txtSellExcl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If txtSellIncl.Text <> "N/A" Then
                txtSellExcl.Text = Format(txtSellIncl.Text / ((100 + Tax) / 100), "0.00")
            End If
            txtSellExcl.Tag = ""
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                txtMarkup.Tag = ""
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
                txtGross.Tag = ""
            Else
                txtGross.Tag = "1"
                txtGross.Text = "0"
                txtGross.Tag = ""
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub CheckforSave()
    If txtProductCode.Text <> "" And txtDescription.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub
Public Sub chkRecipe_Click()
    Select Case chkRecipe.Value
        Case 0
            cmdTab(1).Enabled = False
            txtLandCost.Locked = False
            txtLandCost.Enabled = True
            lblCost.Caption = "Landed Cost:"
        Case 1
            cmdTab(1).Enabled = True
            Select Case cmbUnit.Text
                Case "Preparation Recipe"
                    txtLandCost.Enabled = False
                    txtLandCost.Text = "N/A"
                Case Else
                    txtLandCost.Locked = True
                    lblCost.Caption = "Theoretical Cost:"
            End Select
    End Select
End Sub

Public Sub chkSales_Click()
    Select Case chkSales.Value
        Case 0
            chkTouch.Enabled = False
            chkTouch.Value = 0
            txtSellExcl.Text = "N/A"
            txtSellIncl.Text = "N/A"
            txtMarkup.Text = "N/A"
            txtGross.Text = "N/A"
            txtMarkup.Enabled = False
            txtGross.Enabled = False
            txtSellExcl.Enabled = False
            txtSellIncl.Enabled = False
            cmbTax.Enabled = False
            ButtonEx3.Enabled = False
        Case 1
            If txtSellIncl.Tag = "1" Then Exit Sub
            chkTouch.Enabled = True
            txtSellExcl.Enabled = True
            txtSellIncl.Enabled = True
            txtMarkup.Enabled = True
            txtGross.Enabled = True
            txtMarkup.Text = "0"
            txtGross.Text = "0"
            txtSellExcl.Text = "0.00"
            txtSellIncl.Text = "0.00"
            cmbTax.Enabled = True
            ButtonEx3.Enabled = True
    End Select
End Sub

Public Sub chkScale_Click()
    Select Case chkScale.Value
        Case 0
            cmbScalePrefix.Enabled = False
            lblScale.Enabled = False
            cmbembedtype.Enabled = False
            lblembedtype.Enabled = False
            cmbScalePrefix.BorderColor = &HC0C0C0
        Case 1
            cmbScalePrefix.Enabled = True
            lblScale.Enabled = True
            lblembedtype.Enabled = True
            cmbembedtype.Enabled = True
            cmbScalePrefix.BorderColor = &H0&
    End Select
End Sub

Public Sub chkStock_Click()
    Select Case chkStock.Value
        Case 0
            chkRecipe.Enabled = True
            chkDelete.Enabled = False
            chkDelete.Value = 0
            chkProduction.Enabled = False
            chkDeposit.Value = 0
            chkProduction.Value = 0
        Case 1
            chkRecipe.Enabled = False
            chkRecipe.Value = 0
            chkDelete.Enabled = True
            chkProduction.Value = 0
            chkProduction.Enabled = True
    End Select
End Sub

Private Sub cmbDepart_Change()
    On Error Resume Next
    If cmbDepart.Tag = "1" Then
        Select Case cmbDepart.Text
                Case "<All Departments>"
                    ProductFilter(1) = ""
                Case Else
                    If cmbDepart.Text = "<Unbound>" Then
                        ProductFilter(1) = "Department_No is Null"
                    Else
                        If InStr(Mid(cmbDepart.Text, 1, InStr(cmbDepart, " -") - 1), "-") <> 0 Then
                            ProductFilter(1) = "Department_No = '" & Mid(cmbDepart.Text, 1, InStr(cmbDepart, " -") - 1) & "'"
                        Else
                            ProductFilter(1) = "Left(Department_No," & Len(Mid(cmbDepart.Text, 1, InStr(cmbDepart, " -") - 1)) + 1 & ") ='" & Mid(cmbDepart.Text, 1, InStr(cmbDepart, " -") - 1) & "-'"
                        End If
                    End If
        End Select
        Load_Products
    End If
    On Error GoTo 0
End Sub

Private Sub cmbDepart_GotFocus()
    cmbDepart.Tag = "1"
End Sub

Private Sub cmbDepart_LostFocus()
    cmbDepart.Tag = ""
End Sub

Private Sub cmbDepartments_Change()
    If txtProductCode = "" Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
    End If
    If picDepartments.Visible = True And cmbDepartments.Text <> "<Unbound>" Then
        grdMajor1.Tag = "Not"
        If grdMajor1.FindRow(Trim(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, "-") - 1)), 0, 0, 1, 1) <> -1 Then
            grdMajor1.Row = grdMajor1.FindRow(Trim(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, "-") - 1)), 0, 0, 1, 1)
        End If
        If grdSub1.FindRow(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, " -") - 1), 0, 0, 1, 1) <> -1 Then
            grdSub1.Row = grdSub1.FindRow(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, " -") - 1), 0, 0, 1, 1)
        End If
        grdMajor1.Tag = ""
    End If
    If cmbDepartments.Text = "<Unbound>" Or cmbDepartments.Text = "" Then '
        lblRange.Caption = "0 to 0"
    Else
        ActiveReadServer1 "Select Range_Start,Range_Stop from Departments where Department_No='" & Trim(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, "-") - 1)) & "'"
        If rs1.RecordCount > 0 Then
            lblRange.Caption = rs1.Fields("Range_Start") & " to " & rs1.Fields("Range_Stop")
        Else
             lblRange.Caption = "0 to 0"
        End If
        rs1.Close
    End If
End Sub

Private Sub cmbDepartments_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbDepartments_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
                If txtProductCode.Text = "" Then txtProductCode.SetFocus
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                If txtUnitSize.Enabled = True Then
                    txtUnitSize.SetFocus
                Else
                    cmbUnit.SetFocus
                End If
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
End Sub

Private Sub cmbTax_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbTax_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub

Private Sub cmbTax_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtSellExcl.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
End Sub
Private Sub cmbType_Change()
    Screen.MousePointer = 11
    DoEvents
    If cmbType.Tag = "1" Then
        Select Case cmbType.Text
                Case "<All Products>"
                    ProductFilter(0) = ""
                Case "Stock Keeping Item"
                    ProductFilter(0) = "Stock_Item = 1"
                Case "Sales Item"
                    ProductFilter(0) = "Sales_Item = 1"
                Case "Recipe Item"
                    ProductFilter(0) = "Recipe_Item = 1"
                Case "Preparation Recipes"
                    ProductFilter(0) = "Unit_of_Measure = 'Preparation Recipe'"
                Case "Display on Touch"
                    ProductFilter(0) = "Touch_Item = 1"
                Case "Deposit Bearing Item"
                    ProductFilter(0) = "Returnable_Item = 1"
                Case "Scale Item"
                    ProductFilter(0) = "Scale_Item = 1"
                Case "Delete when Stock Runs Out"
                    ProductFilter(0) = "Once_off = 1"
                Case "Production Item"
                    ProductFilter(0) = "Production_Item = 1"
'                Case "Stock Levels Low"
'                    ProductFilter(3) = "Stock_Level_Min = 1"
'Productsmorne
        End Select
        Load_Products
    End If
    Screen.MousePointer = 0
End Sub
Private Sub cmbType_GotFocus()
    cmbType.Tag = "1"
End Sub

Private Sub cmbType_LostFocus()
    cmbType.Tag = ""
End Sub

Private Sub cmbUnit_Change()
    If cmbUnit.Text = "each" Or cmbUnit.Text = "Preparation Recipe" Then
        txtUnitSize.Text = ""
        txtUnitSize.Enabled = False
    Else
        txtUnitSize.Enabled = True
        chkStock.Enabled = True
        chkSales.Enabled = True
    End If
    If cmbUnit.Text = "Preparation Recipe" Then
        chkStock.Value = 0
        chkSales.Value = 0
        chkStock.Enabled = False
        chkSales.Enabled = False
    Else
        chkStock.Enabled = True
        chkSales.Enabled = True
    End If
End Sub
Private Sub cmbUnit_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbUnit_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub
Private Sub cmbUnit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtShort.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
End Sub
Private Sub cmdBarcode_Click()
    On Error Resume Next
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
    frmProducts.Tag = "Not Now"
    Load frmProdChange
    frmProdChange.Tag = "Products"
    frmProdChange.Show vbModal
    ActiveReadServer "Select * from Products where Product_Code='" & grdProd.TextMatrix(grdProd.Row, 0) & "'"
    If rs.RecordCount > 0 Then
        txtProductCode.Text = rs.Fields("Product_Code")
        txtDescription.Text = rs.Fields("Description")
        txtPackSize.Text = rs.Fields("Pack_Size")
        txtShort.Text = rs.Fields("Short_Description")
        If rs.Fields("Unit_of_Measure") & "" = "" Then
            cmbUnit.Text = "each"
        Else
            cmbUnit.Text = rs.Fields("Unit_of_Measure") & ""
        End If
        If rs.Fields("Unit_Size") = 0 Then
            txtUnitSize.Text = ""
        Else
            txtUnitSize.Text = rs.Fields("Unit_Size") & ""
        End If
        cmbDepartments.Text = grdProd.TextMatrix(grdProd.Row, 2)
        txtLandCost.Text = Format(grdProd.TextMatrix(grdProd.Row, 4), "0.00")
        txtLandCost.ToolTipText = " Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & " "
        cmbTax.Text = grdProd.TextMatrix(grdProd.Row, 5)
        txtSellIncl.Text = Format(grdProd.TextMatrix(grdProd.Row, 6), "0.00")
        If rs.Fields("Scale_Prefix") & "" = "" Then
            cmbScalePrefix.Text = "20"
        Else
            cmbScalePrefix.Text = rs.Fields("Scale_Prefix") & ""
        End If
        If rs.Fields("Kitchen1") & "" = "" Then
            cmbPrinter1.Text = "<None>"
        Else
            cmbPrinter1.Text = rs.Fields("Kitchen1")
        End If
        If rs.Fields("Kitchen2") & "" = "" Then
            cmbPrinter2.Text = "<None>"
        Else
            cmbPrinter2.Text = rs.Fields("Kitchen2")
        End If
        '***************************
        txtSellIncl.Tag = "1"
        chkStock.Value = rs.Fields("Stock_Item")
        chkSales.Value = rs.Fields("Sales_Item")
        chkRecipe.Value = rs.Fields("Recipe_Item")
        chkTouch.Value = rs.Fields("Touch_Item")
        txtEmpty.Text = Val(rs.Fields("Weight_Empty") & "")
        txtFull.Text = Val(rs.Fields("Weight_Full") & "")
        chkDeposit.Value = rs.Fields("Returnable_Item")
        chkScale.Value = rs.Fields("Scale_Item")
        chkDelete.Value = rs.Fields("Once_Off")
        chkWhole.Value = Val(rs.Fields("Whole_Unit") & "")
        If chkSales.Value = 1 Then
            txtSellExcl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If txtSellIncl.Text <> "N/A" Then
                txtSellExcl.Text = Format(txtSellIncl.Text / ((100 + Tax) / 100), "0.00")
            End If
            txtSellExcl.Tag = ""
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                txtMarkup.Tag = ""
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
                txtGross.Tag = ""
            Else
                txtGross.Tag = "1"
                txtGross.Text = "0"
                txtGross.Tag = ""
            End If
            chkTouch.Enabled = True
            txtSellExcl.Enabled = True
            txtSellIncl.Enabled = True
            txtMarkup.Enabled = True
            txtGross.Enabled = True
            cmbTax.Enabled = True
            ButtonEx3.Enabled = True
        End If
        txtSellIncl.Tag = ""
    End If
    rs.Close
    On Error Resume Next
End Sub
Private Sub cmbPrinter1_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbPrinter2_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbScalePrefix_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub Form_Load()
    Dim x As Printer
    SafetyCode "SOH"
    SafetyCode "Recipe"
    adoData.ConnectionString = cnnMain.ConnectionString
    adoData.CursorLocation = adUseServer
    adoData.CursorType = adOpenStatic
    adoData.LockType = adLockReadOnly
    
    adoData.RecordSource = "Select * from Product_List " & lblOrderFilter.Caption
    adoData.Refresh
    Screen.MousePointer = 11
    frmMain.cmdMenu(0).Value = 1
    RecordCount = 0
    ProductFilter(0) = ""
    ProductFilter(1) = ""
    ProductFilter(2) = ""
    grdSuppliers.Rows = 1
    grdSuppliers.TextMatrix(0, 0) = "Supplier No."
    grdSuppliers.TextMatrix(0, 1) = "Supplier Name"
    grdSuppliers.TextMatrix(0, 2) = "Qty Received"
    grdSuppliers.TextMatrix(0, 3) = "Delivered On"
    grdSuppliers.TextMatrix(0, 4) = "Landed Cost"
    grdSuppliers.ColWidth(0) = grdSuppliers.Width * 0.15
    grdSuppliers.ColWidth(1) = grdSuppliers.Width * 0.35
    grdSuppliers.ColWidth(2) = grdSuppliers.Width * 0.15
    grdSuppliers.ColAlignment(2) = flexAlignRightCenter
    grdSuppliers.ColWidth(3) = grdSuppliers.Width * 0.2
    grdSuppliers.ColWidth(4) = grdSuppliers.Width * 0.12
    grdSuppliers.ColAlignment(4) = flexAlignRightCenter
    
    grdRecipe.Rows = 1
    grdRecipe.TextMatrix(0, 0) = "Recipe Line"
    grdRecipe.TextMatrix(0, 1) = "Description or Product"
    grdRecipe.TextMatrix(0, 2) = "Unit of Measure"
    grdRecipe.TextMatrix(0, 3) = "Qty Used"
    grdRecipe.TextMatrix(0, 4) = "Landed Cost"
    grdRecipe.ColWidth(0) = grdRecipe.Width * 0.2
    grdRecipe.ColWidth(1) = grdRecipe.Width * 0.35
    grdRecipe.ColWidth(2) = grdRecipe.Width * 0.15
    grdRecipe.ColAlignment(0) = flexAlignLeftCenter
    grdRecipe.ColAlignment(1) = flexAlignLeftCenter
    grdRecipe.ColAlignment(3) = flexAlignRightCenter
    grdRecipe.ColWidth(3) = grdRecipe.Width * 0.15
    grdRecipe.ColWidth(4) = grdRecipe.Width * 0.12
    grdRecipe.ColAlignment(4) = flexAlignRightCenter
    
    frmMain.Toolbar1.Buttons(5).Enabled = False
    
    grdProd.TextMatrix(0, 0) = "Product Code"
    grdProd.TextMatrix(0, 1) = "Description"
    grdProd.TextMatrix(0, 2) = "Department"
    grdProd.TextMatrix(0, 3) = "Stock on Hand "
    grdProd.TextMatrix(0, 4) = "Landed Cost "
    grdProd.TextMatrix(0, 5) = "Tax Rate "
    grdProd.TextMatrix(0, 6) = "Price (Incl) "
    grdProd.ColAlignment(0) = flexAlignLeftCenter
    grdProd.ColAlignment(1) = flexAlignLeftCenter
    grdProd.ColAlignment(2) = flexAlignLeftCenter
    grdProd.ColAlignment(3) = flexAlignRightCenter
    grdProd.ColAlignment(4) = flexAlignRightCenter
    grdProd.ColAlignment(5) = flexAlignRightCenter
    grdProd.ColAlignment(6) = flexAlignRightCenter
    grdProd.ColWidth(0) = grdProd.Width * 0.12
    grdProd.ColWidth(1) = grdProd.Width * 0.3
    grdProd.ColWidth(2) = grdProd.Width * 0.15
    grdProd.ColWidth(3) = grdProd.Width * 0.1
    grdProd.ColWidth(4) = grdProd.Width * 0.1
    grdProd.ColWidth(5) = grdProd.Width * 0.1
    grdProd.ColWidth(6) = grdProd.Width * 0.1
    grdProd.ColFormat(4) = "0.00"
    grdProd.ColFormat(6) = "0.00"
    For i = 7 To grdProd.Cols - 1
        grdProd.ColHidden(i) = True
    Next i
    grdRecipe.ColHidden(5) = True
    grdRecipe.ColHidden(6) = True
    grdRecipe.ColHidden(7) = True
    
    grdPrices.RowHeight(0) = 400
    grdPrices.RowHeight(1) = 400
    For i = 0 To 5
        If i > 0 Then
            If Branch_Type = 10 Then
                grdPrices.TextMatrix(1, i) = "Retail " & i
                grdPrices.TextMatrix(0, i) = "  Selling Prices"
            Else
                grdPrices.TextMatrix(1, i) = "Price " & i + 1
                grdPrices.TextMatrix(0, i) = "  Selling Prices"
            End If
        End If
        grdPrices.Cell(flexcpAlignment, 0, i, 0, i) = flexAlignLeftCenter
        grdPrices.Cell(flexcpAlignment, 1, i, 1, i) = flexAlignCenterCenter
        grdPrices.Cell(flexcpFontSize, 0, i, 0, i) = 10
    Next i
    grdPrices.MergeRow(0) = True
    grdPrices.TextMatrix(2, 0) = "Landed Cost:"
    grdPrices.TextMatrix(3, 0) = "Markup%:"
    grdPrices.TextMatrix(4, 0) = "Gross Profit%:"
    grdPrices.TextMatrix(5, 0) = "Selling Price (excl):"
    grdPrices.TextMatrix(6, 0) = "Tax:"
    grdPrices.TextMatrix(7, 0) = "Selling Price (incl):"
    grdRecipe.Rows = 2
    For i = 2 To 7
        grdPrices.RowHeight(i) = 300
        grdPrices.Cell(flexcpAlignment, i, 0, i, 0) = flexAlignRightCenter
        grdPrices.Cell(flexcpAlignment, i, 1, i, 5) = flexAlignLeftCenter
    Next i
    grdPrices.Cell(flexcpBackColor, 5, 1, 5, 5) = &HF9E6D7
    grdPrices.Cell(flexcpBackColor, 7, 1, 7, 5) = &HF9E6D7
    
    For i = 1 To 3
        picBox(i).Visible = False
    Next i
    cmbScalePrefix.Clear
    cmbScalePrefix.AddItem "<None>"
    For i = 0 To 9
        cmbScalePrefix.AddItem Trim("2" & Trim(Str(i)))
    Next i
    
    cmbScalePrefix.Text = "<None>"
    
    cmbembedtype.AddItem "<None>"
     cmbembedtype.AddItem "Price"
     cmbembedtype.AddItem "Weight"
    cmbembedtype.Text = "<None>"
    
    
    
    LoadUnitofMeasure
    txtFull.Text = 0
    txtEmpty.Text = 0
    cmbType.Clear
        
    cmbType.AddItem "<All Products>"
    cmbType.AddItem "Stock Keeping Item"
    cmbType.AddItem "Sales Item"
    cmbType.AddItem "Recipe Item"
    cmbType.AddItem "Preparation Recipes"
    cmbType.AddItem "Display on Touch"
    cmbType.AddItem "Deposit Bearing Item"
    cmbType.AddItem "Scale Item"
    cmbType.AddItem "Delete when Stock Runs Out"
    cmbType.AddItem "Production Item"
    cmbType.Text = "<All Products>"
    cmbPrinter1.Clear
    cmbPrinter2.Clear
    cmbPrinter1.AddItem "<None>"
    cmbPrinter2.AddItem "<None>"
    For Each x In Printers
        cmbPrinter1.AddItem x.DeviceName
        cmbPrinter2.AddItem x.DeviceName
    Next
    cmbPrinter1.Text = "<None>"
    cmbPrinter2.Text = "<None>"
    cmbTax.Clear
    ActiveReadServer "Select * from Tax_Rates order by Tax_Type"
    While Not rs.EOF
        cmbTax.AddItem rs.Fields("Tax_type") & " - " & rs.Fields("Tax_Rate") & "% " & rs.Fields("Description")
        rs.MoveNext
    Wend
    rs.Close
    If cmbTax.ListCount > 1 Then
        cmbTax.Text = cmbTax.List(0)
    End If
    
    cmbDepartments.Clear
    cmbDepart.Clear
    
    ActiveReadServer "Select * from Departments order by Department_No"
    cmbDepartments.AddItem "<Unbound>"
    cmbDepart.AddItem "<All Departments>"
    cmbDepart.AddItem "<Unbound>"
    While Not rs.EOF
        cmbDepartments.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        cmbDepart.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmbDepartments.Text = "<Unbound>"
    cmbDepart.Text = "<All Departments>"
    For i = 1 To 5
        grdPrices.TextMatrix(2, i) = txtLandCost.Text
        grdPrices.TextMatrix(3, i) = "0"
        grdPrices.TextMatrix(4, i) = "0"
        grdPrices.TextMatrix(5, i) = "0.00"
        If cmbTax.Text <> "" Then
            grdPrices.TextMatrix(6, i) = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
        End If
        grdPrices.TextMatrix(7, i) = "0.00"
    Next i
    grdProd.Row = 0
    If grdProd.Rows > 2 Then
        grdProd.Row = 1
        frmMain.Toolbar1.Buttons(4).Enabled = True
    End If
    RecordCount = adoData.Recordset.RecordCount
    
    frmMain.cmdMenu(0).Value = 1
    frmMain.stbBar.Panels(3) = "Records = " & Val(RecordCount)
    Screen.MousePointer = 0
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub
Public Sub Load_Products()
    Screen.MousePointer = 11
    Dim x As Printer
    RecordCount = 0
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
    cmbPrinter1.Clear
    cmbPrinter2.Clear
    cmbPrinter1.AddItem "<None>"
    cmbPrinter2.AddItem "<None>"
    For Each x In Printers
        cmbPrinter1.AddItem x.DeviceName
        cmbPrinter2.AddItem x.DeviceName
    Next
    cmbPrinter1.Text = "<None>"
    cmbPrinter2.Text = "<None>"
    
    cmbTax.Clear
    ActiveReadServer "Select * from Tax_Rates order by Tax_Type"
    While Not rs.EOF
        cmbTax.AddItem rs.Fields("Tax_type") & " - " & rs.Fields("Tax_Rate") & "% " & rs.Fields("Description")
        rs.MoveNext
    Wend
    rs.Close
    If cmbTax.ListCount > 1 Then
        cmbTax.Text = cmbTax.List(0)
    End If
    
    cmbDepartments.Clear
    ActiveReadServer "Select * from Departments order by Department_No"
    cmbDepartments.AddItem "<Unbound>"
    While Not rs.EOF
        cmbDepartments.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmbDepartments.Text = "<Unbound>"
    DoEvents
    txtProductCode.Text = ""
    grdProd.Rows = 1
    
    DoEvents
    adoData.ConnectionString = cnnMain.ConnectionString
    adoData.CursorLocation = adUseClient
    adoData.CursorType = adOpenKeyset
    adoData.LockType = adLockBatchOptimistic
    
    If Trim(ProductFilter(0)) = "" And Trim(ProductFilter(1)) = "" Then
        adoData.RecordSource = "Select * from Product_List order by Description"
    End If
    
    If Trim(ProductFilter(0)) <> "" Then
        If Trim(ProductFilter(1)) = "" Then
            adoData.RecordSource = "Select * from Product_List where " & Trim(ProductFilter(0)) & " order by Description"
        End If
    End If
    If Trim(ProductFilter(1)) <> "" Then
        If Trim(ProductFilter(0)) <> "" Then
            adoData.RecordSource = "Select * from Product_List where " & Trim(ProductFilter(0)) & " and " & Trim(ProductFilter(1)) & " order by Description"
        Else
            adoData.RecordSource = "Select * from Product_List where " & Trim(ProductFilter(1)) & " order by Description"
        End If
    End If
    If Trim(ProductFilter(2)) <> "" Then
        If InStr(adoData.RecordSource, "Product_List where") = 0 Then
            adoData.RecordSource = Replace(adoData.RecordSource, " order by Description", "") & " where Description like '%" & Trim(ProductFilter(2)) & "%' order by Description"
        Else
            adoData.RecordSource = Replace(adoData.RecordSource, " order by Description", "") & " and Description like '%" & Trim(ProductFilter(2)) & "%' order by Description"
        End If
        ProductFilter(2) = ""
    End If
  
    
    
    
    
'            If rs.Fields("Stock_Level_Min") < 10 Then
'
'            grdProd.Cell(flexcpBackColor, 1, 0) = &HC0C0FF
'            grdProd.Cell(flexcpBackColor, 1, 1) = &HC0C0FF
'            grdProd.Cell(flexcpBackColor, 1, 2) = &HC0C0FF
'            grdProd.Cell(flexcpBackColor, 1, 3) = &HC0C0FF
'            grdProd.Cell(flexcpBackColor, 1, 4) = &HC0C0FF
'            grdProd.Cell(flexcpBackColor, 1, 5) = &HC0C0FF
'            grdProd.Cell(flexcpBackColor, 1, 6) = &HC0C0FF
'            End If
    
    
    
    
    
    
    
    
    adoData.Refresh
    RecordCount = adoData.Recordset.RecordCount
    grdProd.Row = 0
    grdProd.TextMatrix(0, 0) = "Product Code"
    grdProd.TextMatrix(0, 1) = "Description"
    grdProd.TextMatrix(0, 2) = "Department"
    grdProd.TextMatrix(0, 3) = "Stock on Hand "
    grdProd.TextMatrix(0, 4) = "Landed Cost "
    grdProd.TextMatrix(0, 5) = "Tax Rate "
    grdProd.TextMatrix(0, 6) = "Price (Incl) "
    grdProd.ColAlignment(0) = flexAlignLeftCenter
    grdProd.ColAlignment(1) = flexAlignLeftCenter
    grdProd.ColAlignment(2) = flexAlignLeftCenter
    grdProd.ColAlignment(3) = flexAlignRightCenter
    grdProd.ColAlignment(4) = flexAlignRightCenter
    grdProd.ColAlignment(5) = flexAlignRightCenter
    grdProd.ColAlignment(6) = flexAlignRightCenter
    grdProd.ColWidth(0) = grdProd.Width * 0.12
    grdProd.ColWidth(1) = grdProd.Width * 0.3
    grdProd.ColWidth(2) = grdProd.Width * 0.15
    grdProd.ColWidth(3) = grdProd.Width * 0.1
    grdProd.ColWidth(4) = grdProd.Width * 0.1
    grdProd.ColWidth(5) = grdProd.Width * 0.1
    grdProd.ColWidth(6) = grdProd.Width * 0.1
    grdProd.ColFormat(4) = "0.00"
    grdProd.ColFormat(6) = "0.00"
    For i = 7 To grdProd.Cols - 1
        grdProd.ColHidden(i) = True
    Next i
    DoEvents
    If grdProd.Rows = 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
        txtDescription.Text = ""
        txtPackSize.Text = "1"
        txtShort.Text = ""
        cmbUnit.Text = "each"
        txtUnitSize.Text = ""
        txtLandCost.Text = "0.00"
        txtMarkup.Text = "0"
        txtGross.Text = "0"
        txtSellExcl.Text = "0.00"
        If cmbTax.ListCount > 0 Then cmbTax.Text = cmbTax.List(0)
        txtSellIncl.Text = "0.00"
        
        
    Else
        grdProd.Row = 1
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
        ActiveReadServer "select * from Product_Prices where Product_Code= '" & txtProductCode.Text & "'"
        If rs.RecordCount > 0 Then
            grdPrices.TextMatrix(7, 1) = Format(rs.Fields("Price2"), "0.00")
            grdPrices.TextMatrix(7, 2) = Format(rs.Fields("Price3"), "0.00")
            grdPrices.TextMatrix(7, 3) = Format(rs.Fields("Price4"), "0.00")
            grdPrices.TextMatrix(7, 4) = Format(rs.Fields("Price5"), "0.00")
            grdPrices.TextMatrix(7, 5) = Format(rs.Fields("Price6"), "0.00")
        End If
        rs.Close
    End If
    frmMain.stbBar.Panels(3) = "Records = " & Val(RecordCount)
    txtProductCode.SetFocus
    Screen.MousePointer = 0
End Sub

Private Sub LoadUnitofMeasure()
    cmbUnit.Clear
    cmbUnit.AddItem "ml"
    cmbUnit.AddItem "lt"
    cmbUnit.AddItem "g"
    cmbUnit.AddItem "kg"
    cmbUnit.AddItem "ton"
    cmbUnit.AddItem "each"
    cmbUnit.AddItem "box"
    cmbUnit.AddItem "Preparation Recipe"
    cmbUnit.Text = "each"
End Sub
Public Sub CreateProduct()
    On Error GoTo far
    Screen.MousePointer = 11
    DoEvents
    If cmdUp.Caption = "6" Then
         picTopFrame.top = 5670
         picSeperate.top = 5580
         grdProd.Height = 2640
         grdProd.top = 6240
         cmdUp.Caption = "5"
     End If
     If cmbDepartments.Text = "<Unbound>" Then
         txtProductCode.Text = ""
         txtProductCode.SetFocus
         chkStock.Value = 1
         chkSales.Value = 1
         chkProduction.Value = 0
     Else
         If Len(txtProductCode.Text) = 2 And Asc(Left(txtProductCode.Text, 1)) > 60 Then
            ActiveReadServer "Select Top 1 Product_Code from Products where Product_Code like '" & txtProductCode.Text & "%' order by Product_Code Desc"
            If rs.RecordCount > 0 Then
                txtProductCode.Text = txtProductCode.Text & Format(Trim(Str(Val(Mid(rs.Fields("Product_Code"), Len(txtProductCode.Text) + 1)) + 1)), "000")
            Else
                txtProductCode.Text = ""
            End If
            ActiveReadServer "SELECT Locations.Loc_Type " & _
            "FROM Department_Links INNER JOIN " & _
            "Locations ON Department_Links.Location_No = Locations.Location_No " & _
            "WHERE (Department_Links.Dept_No = '" & Trim(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, " -") - 1)) & "')"
            While Not rs.EOF
                 If rs.Fields("Loc_Type") = 0 Then
                     chkSales.Value = 1
                 End If
                 If rs.Fields("Loc_Type") = 1 Then
                     chkStock.Value = 1
                 End If
                 rs.MoveNext
             Wend
             rs.Close
             txtDescription.SetFocus
         Else
             If lblRange.Caption = "0 to 0" Then
                 MsgBox "HeroPOS cannot provide you with a Product Code for this Department" & Chr$(13) & "as you have not set up a Start and Stop range on the Department.", vbCritical, "HeroPOS"
             Else
                FirstCode = 0
                ActiveReadServer1 "Select Range_Start,Range_Stop from Departments where Department_No='" & Trim(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, "-") - 1)) & "'"
                If rs1.RecordCount > 0 Then
                    FirstCode = rs1.Fields("Range_Start")
                    LastCode = rs1.Fields("Range_Stop")
                End If
                rs1.Close
                If FirstCode = 0 Then FirstCode = FirstCode + 1
                For i = FirstCode To LastCode
                    ActiveReadServer1 "Select * from Products where Product_Code = '" & i & "'"
                    If rs1.RecordCount = 0 Then
                        rs1.Close
                        txtProductCode.Text = i
                        GoTo far1
                    End If
                Next i
                rs1.Close
                
                MsgBox "Your maximum code range for this department was exceeded." & Chr(13) & "Please select a higher range.", vbCritical
                Screen.MousePointer = 0
                On Error GoTo 0
                Exit Sub
             End If
far1:
            ActiveReadServer "SELECT Locations.Loc_Type " & _
            "FROM Department_Links INNER JOIN " & _
            "Locations ON Department_Links.Location_No = Locations.Location_No " & _
            "WHERE (Department_Links.Dept_No = '" & Trim(Mid(cmbDepartments.Text, 1, InStr(cmbDepartments.Text, " -") - 1)) & "')"
            While Not rs.EOF
                 If rs.Fields("Loc_Type") = 0 Then
                     chkSales.Value = 1
                 End If
                 If rs.Fields("Loc_Type") = 1 Then
                     chkStock.Value = 1
                 End If
                 rs.MoveNext
             Wend
             rs.Close
             txtDescription.SetFocus
        End If
     End If
     txtDescription.Text = ""
     txtPackSize.Text = "1"
     txtShort.Text = ""
     cmbUnit.Text = "each"
     txtUnitSize.Text = ""
     txtLandCost.Text = "0.00"
     txtMarkup.Text = "0"
     txtGross.Text = "0"
     txtSellExcl.Text = "0.00"
     txtEmpty.Text = "0"
     txtFull.Text = "0"
     chkProduction.Value = 0
     If cmbTax.ListCount > 0 Then cmbTax.Text = cmbTax.List(0)
     txtSellIncl.Text = "0.00"
     For i = 1 To 5
         grdPrices.TextMatrix(2, i) = txtLandCost.Text
         grdPrices.TextMatrix(3, i) = "0"
         grdPrices.TextMatrix(4, i) = "0"
         grdPrices.TextMatrix(5, i) = "0.00"
         grdPrices.TextMatrix(6, i) = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
         grdPrices.TextMatrix(7, i) = "0.00"
     Next i
    grdRecipe.Rows = 1
    grdRecipe.Rows = 2
    cmdTab(0).Value = 1
    picBox(0).Visible = True
    cmdTestR.Visible = False
    cmdPrep.Visible = False
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(4).Enabled = False
    frmMain.Toolbar1.Buttons(5).Enabled = False
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
far:
    Resume Next
End Sub
Private Sub Load_Recipe()
    grdRecipe.ColComboList(1) = ""
End Sub
Private Sub txtProductCode_Change()
    TopRow = 0
    BottomRow = 0
    For i = 1 To Len(txtProductCode.Text) - 1
        If Len(txtProductCode.Text) < 9 Then
            If i / 2 <> Int(i / 2) Then
                TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
            End If
        Else
            If i / 2 = Int(i / 2) Then
                TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
            End If
        End If
    Next i
    TopRow = TopRow * 3
    Result = TopRow + BottomRow
    Result = (1 - ((Result / 10) - Int((Result / 10)))) * 10
    If Result = 10 Then Result = 0
    If Round(Result, 0) = Int(Val(Right(txtProductCode, 1))) Then
        PicBC.Visible = True
    Else
        PicBC.Visible = False
    End If
    CheckforSave
    If grdProd.FindRow(txtProductCode.Text, 0, 0, 0, 1) > 0 And txtDescription <> "" Then
        grdProd.Row = grdProd.FindRow(txtProductCode.Text, 0, 0, 0, 1)
        grdProd.ShowCell grdProd.Row, 0
    Else
        txtDescription.Text = ""
        txtPackSize.Text = "1"
        txtShort.Text = ""
        cmbUnit.Text = "each"
        txtUnitSize.Text = ""
        txtLandCost.Text = "0.00"
        txtMarkup.Text = "0"
        txtGross.Text = "0"
        txtSellExcl.Text = "0.00"
        txtEmpty.Text = "0"
        txtFull.Text = "0"
        If cmbTax.ListCount > 0 Then cmbTax.Text = cmbTax.List(0)
        txtSellIncl.Text = "0.00"
        For i = 1 To 5
            grdPrices.TextMatrix(2, i) = txtLandCost.Text
            grdPrices.TextMatrix(3, i) = "0"
            grdPrices.TextMatrix(4, i) = "0"
            grdPrices.TextMatrix(5, i) = "0.00"
            grdPrices.TextMatrix(6, i) = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            grdPrices.TextMatrix(7, i) = "0.00"
        Next i
        grdRecipe.Rows = 1
        grdRecipe.Rows = 2
        cmdTab(0).Value = 1
        picBox(0).Visible = True
        cmdTestR.Visible = False
        cmdPrep.Visible = False
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
    End If
End Sub
Private Sub txtProductCode_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub
Private Sub txtproductcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            If txtSellIncl.Enabled = True Then
                txtSellIncl.SetFocus
            Else
                txtLandCost.SetFocus
            End If
        Case 40
            txtDescription.SetFocus
    End Select
End Sub
Private Sub txtproductcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case 39
            KeyAscii = 0
        Case 65 To 90
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtSellExcl_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub

Private Sub txtSellExcl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtGross.SetFocus
        Case 40
            cmbTax.SetFocus
    End Select
End Sub

Private Sub txtSellExcl_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtSellExcl_LostFocus()
    If txtSellExcl.Text = "" Then txtSellExcl.Text = "0.00"
    txtSellExcl.Text = Format(txtSellExcl.Text, "0.00")
End Sub

Private Sub txtSellIncl_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub

Private Sub txtSellIncl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            cmbTax.SetFocus
        Case 40
            txtProductCode.SetFocus
    End Select
End Sub

Private Sub txtSellIncl_KeyPress(KeyAscii As Integer)
    If InStr(txtSellIncl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtSellIncl_LostFocus()
    If txtSellIncl.Text = "" Then txtSellIncl.Text = "0.00"
    txtSellIncl.Text = Format(txtSellIncl.Text, "0.00")
End Sub
Private Sub txtMarkup_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub
Private Sub txtMarkup_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtLandCost.SetFocus
        Case 40
            txtGross.SetFocus
    End Select
End Sub
Private Sub txtMarkup_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtMarkup_LostFocus()
    If txtMarkup.Text = "" Then txtMarkup.Text = "0"
End Sub
Private Sub grdMajor1_RowColChange()
    grdSub1.Rows = 1
    ActiveReadServer1 "Select * From Departments where Dept_Type=1 and Dept_Parent='" & grdMajor1.TextMatrix(grdMajor1.Row, 0) & "'"
    i = 0
    While Not rs1.EOF
        i = i + 1
        grdSub1.Rows = grdSub1.Rows + 1
        grdSub1.TextMatrix(i, 0) = rs1.Fields("Department_No")
        grdSub1.TextMatrix(i, 1) = rs1.Fields("Dept_Name")
        rs1.MoveNext
    Wend
    rs1.Close
    For i = 1 To grdMinor1.Rows - 1
        grdMinor1.TextMatrix(i, 2) = "0"
        grdMinor1.Cell(flexcpFontItalic, i, 0, i, 1) = False
        grdMinor1.Cell(flexcpFontBold, i, 0, i, 1) = False
        grdMinor1.Cell(flexcpBackColor, i, 0, i, 2) = ""
    Next i
    ActiveReadServer2 "Select * from Department_links where Dept_No = '" & grdMajor1.TextMatrix(grdMajor1.Row, 0) & "'"
    While Not rs2.EOF
        For i = 1 To grdMinor1.Rows - 1
            If rs2.Fields("Location_No") = grdMinor1.TextMatrix(i, 0) Then
                grdMinor1.TextMatrix(i, 2) = "1"
                grdMinor1.Cell(flexcpFontItalic, i, 0, i, 1) = True
                grdMinor1.Cell(flexcpFontBold, i, 0, i, 1) = True
                grdMinor1.Cell(flexcpBackColor, i, 0, i, 2) = &HFFC0C0
            End If
        Next i
        rs2.MoveNext
    Wend
    rs2.Close
    If grdMajor1.Rows > 1 Then
        If grdMajor1.Tag = "Not" Then Exit Sub
        cmbDepartments.Text = grdMajor1.TextMatrix(grdMajor1.Row, 0) & " - " & grdMajor1.TextMatrix(grdMajor1.Row, 1)
    End If
End Sub

Private Sub grdPrices_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case grdPrices.TextMatrix(grdPrices.Row, 0)
        Case "Landed Cost:"
            If Val(grdPrices.TextMatrix(2, grdPrices.Col)) <> 0 Then
                grdPrices.TextMatrix(3, grdPrices.Col) = Round(((Val(grdPrices.TextMatrix(5, grdPrices.Col)) - Val(grdPrices.TextMatrix(2, grdPrices.Col))) / Val(grdPrices.TextMatrix(2, grdPrices.Col)) * 100), 3)
            End If
            If Val(grdPrices.TextMatrix(5, grdPrices.Col)) <> 0 Then
                grdPrices.TextMatrix(4, grdPrices.Col) = Round(((Val(grdPrices.TextMatrix(5, grdPrices.Col)) - Val(grdPrices.TextMatrix(2, grdPrices.Col))) / Val(grdPrices.TextMatrix(5, grdPrices.Col)) * 100), 3)
            End If
        Case "Markup%:"
            grdPrices.TextMatrix(5, grdPrices.Col) = Format(grdPrices.TextMatrix(2, grdPrices.Col) * ((100 + Val(grdPrices.TextMatrix(3, grdPrices.Col))) / 100), "0.00")
            If Val(grdPrices.TextMatrix(5, grdPrices.Col)) <> 0 Then
                grdPrices.TextMatrix(4, grdPrices.Col) = Round(((Val(grdPrices.TextMatrix(5, grdPrices.Col)) - Val(grdPrices.TextMatrix(2, grdPrices.Col))) / Val(grdPrices.TextMatrix(5, grdPrices.Col)) * 100), 3)
            Else
                grdPrices.TextMatrix(4, grdPrices.Col) = "0"
            End If
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            grdPrices.TextMatrix(7, grdPrices.Col) = Format(grdPrices.TextMatrix(5, grdPrices.Col) * ((100 + Tax) / 100), "0.00")
        Case "Gross Profit%:"
            grdPrices.TextMatrix(5, grdPrices.Col) = Format(grdPrices.TextMatrix(2, grdPrices.Col) / ((100 - Val(grdPrices.TextMatrix(4, grdPrices.Col))) / 100), "0.00")
            If Val(grdPrices.TextMatrix(2, grdPrices.Col)) <> 0 Then
                grdPrices.TextMatrix(3, grdPrices.Col) = Round(((Val(grdPrices.TextMatrix(5, grdPrices.Col)) - Val(grdPrices.TextMatrix(2, grdPrices.Col))) / Val(grdPrices.TextMatrix(2, grdPrices.Col)) * 100), 3)
            End If
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            grdPrices.TextMatrix(7, grdPrices.Col) = Format(grdPrices.TextMatrix(5, grdPrices.Col) * ((100 + Tax) / 100), "0.00")
        Case "Selling Price (excl):"
            grdPrices.TextMatrix(5, grdPrices.Col) = Format(grdPrices.TextMatrix(5, grdPrices.Col), "0.00")
            If Val(grdPrices.TextMatrix(2, grdPrices.Col)) <> 0 Then
                grdPrices.TextMatrix(3, grdPrices.Col) = Round(((Val(grdPrices.TextMatrix(5, grdPrices.Col)) - Val(grdPrices.TextMatrix(2, grdPrices.Col))) / Val(grdPrices.TextMatrix(2, grdPrices.Col)) * 100), 3)
            End If
            If Val(grdPrices.TextMatrix(5, grdPrices.Col)) <> 0 Then
                grdPrices.TextMatrix(4, grdPrices.Col) = Round(((Val(grdPrices.TextMatrix(5, grdPrices.Col)) - Val(grdPrices.TextMatrix(2, grdPrices.Col))) / Val(grdPrices.TextMatrix(5, grdPrices.Col)) * 100), 3)
            Else
                grdPrices.TextMatrix(4, grdPrices.Col) = "0"
            End If
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If grdPrices.TextMatrix(5, grdPrices.Col) <> "N/A" Then
                grdPrices.TextMatrix(7, grdPrices.Col) = Format(grdPrices.TextMatrix(5, grdPrices.Col) * ((100 + Tax) / 100), "0.00")
            End If
        Case "Selling Price (incl):"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If grdPrices.TextMatrix(7, grdPrices.Col) <> "N/A" Then
                grdPrices.TextMatrix(5, grdPrices.Col) = Format(grdPrices.TextMatrix(7, grdPrices.Col) / ((100 + Tax) / 100), "0.00")
            End If
            If Val(grdPrices.TextMatrix(2, grdPrices.Col)) <> 0 Then
                grdPrices.TextMatrix(3, grdPrices.Col) = Round(((Val(grdPrices.TextMatrix(5, grdPrices.Col)) - Val(grdPrices.TextMatrix(2, grdPrices.Col))) / Val(grdPrices.TextMatrix(2, grdPrices.Col)) * 100), 3)
            End If
            If Val(grdPrices.TextMatrix(5, grdPrices.Col)) <> 0 Then
                grdPrices.TextMatrix(4, grdPrices.Col) = Round(((Val(grdPrices.TextMatrix(5, grdPrices.Col)) - Val(grdPrices.TextMatrix(2, grdPrices.Col))) / Val(grdPrices.TextMatrix(5, grdPrices.Col)) * 100), 3)
            Else
                grdPrices.TextMatrix(4, grdPrices.Col) = "0"
            End If
            grdPrices.TextMatrix(7, grdPrices.Col) = Format(grdPrices.TextMatrix(7, grdPrices.Col), "0.00")
    End Select
End Sub

Private Sub grdPrices_EnterCell()
    Select Case grdPrices.Row
        Case 0, 1, 2, 6
            grdPrices.Editable = flexEDNone
        Case 3, 4, 5, 7
            grdPrices.EditMaxLength = 12
            grdPrices.Editable = flexEDKbdMouse
            grdPrices.EditCell
            grdPrices.EditSelStart = 0
            grdPrices.EditSelLength = Len(grdPrices.Text)
    End Select
End Sub
Private Sub grdPrices_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If grdPrices.EditSelLength = Len(grdPrices.Text) Then
        Select Case KeyCode
            Case 13
                If grdPrices.Row = grdPrices.Rows - 1 Then grdPrices.Row = 3: Exit Sub
                grdPrices.Row = grdPrices.Row + 1
            Case 37
                If grdPrices.Col = 1 Then grdPrices.Col = grdPrices.Cols - 1: Exit Sub
                grdPrices.Col = grdPrices.Col - 1
            Case 39
                If grdPrices.Col = grdPrices.Cols - 1 Then grdPrices.Col = 1: Exit Sub
                grdPrices.Col = grdPrices.Col + 1
        End Select
    End If
End Sub
Private Sub grdPrices_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(grdPrices.EditText, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 13, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub grdProd_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If grdProd.Rows = 1 Or grdProd.Tag = "1" Then Exit Sub
    frmMain.Toolbar1.Buttons(5).Enabled = True
    If OldRow <> NewRow Then
        ActiveReadServer "Select * from Products where Product_Code='" & grdProd.TextMatrix(grdProd.Row, 0) & "'"
        If rs.RecordCount > 0 Then
            txtProductCode.Text = rs.Fields("Product_Code")
            TopRow = 0
            BottomRow = 0
            For i = 1 To Len(txtProductCode.Text) - 1
                If Len(txtProductCode.Text) < 9 Then
                    If i / 2 <> Int(i / 2) Then
                        TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
                    Else
                        BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
                    End If
                Else
                    If i / 2 = Int(i / 2) Then
                        TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
                    Else
                        BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
                    End If
                End If
            Next i
            TopRow = TopRow * 3
            Result = TopRow + BottomRow
            Result = (1 - ((Result / 10) - Int((Result / 10)))) * 10
            If Result = 10 Then Result = 0
            If Round(Result, 0) = Int(Val(Right(txtProductCode, 1))) Then
                PicBC.Visible = True
            Else
                PicBC.Visible = False
            End If
            txtDescription.Text = rs.Fields("Description")
            txtPackSize.Text = rs.Fields("Pack_Size") & ""
            txtShort.Text = rs.Fields("Short_Description")
            txtRef.Text = rs.Fields("Nappi_Code") & ""
            If rs.Fields("Unit_of_Measure") & "" = "" Then
                cmbUnit.Text = "each"
            Else
                cmbUnit.Text = rs.Fields("Unit_of_Measure") & ""
            End If
            If rs.Fields("Unit_Size") = 0 Then
                txtUnitSize.Text = ""
            Else
                txtUnitSize.Text = rs.Fields("Unit_Size") & ""
            End If
            cmbDepartments.Text = grdProd.TextMatrix(grdProd.Row, 2)
            txtLandCost.Text = Format(grdProd.TextMatrix(grdProd.Row, 4), "0.00")
            txtLandCost.ToolTipText = " Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & " "
            cmbTax.Text = grdProd.TextMatrix(grdProd.Row, 5)
            txtSellIncl.Text = Format(grdProd.TextMatrix(grdProd.Row, 6), "0.00")
            If rs.Fields("Scale_Prefix") & "" = "" Then
                cmbScalePrefix.Text = "20"
            Else
                cmbScalePrefix.Text = rs.Fields("Scale_Prefix") & ""
            End If
            
            If rs.Fields("Scaleitemtype") & "" = "" Then
            cmbembedtype.Text = "<None>"
            Else
            cmbembedtype.Text = rs.Fields("Scaleitemtype") & ""
            End If
            
            
            

            
            
            
            If rs.Fields("Stock_Level_Min") & "" = "" Or rs.Fields("Stock_Level_Min") = "0" Then
            Txtstocklevel.Text = "None"
            Else
            Txtstocklevel.Text = rs.Fields("Stock_Level_Min") & ""
            
            

            
     
            End If
            
            
     
            
            
            
            If rs.Fields("Kitchen1") & "" = "" Then
                cmbPrinter1.Text = "<None>"
            Else
                cmbPrinter1.Text = rs.Fields("Kitchen1")
            End If
            If rs.Fields("Kitchen2") & "" = "" Then
                cmbPrinter2.Text = "<None>"
            Else
                cmbPrinter2.Text = rs.Fields("Kitchen2")
            End If
            '***************************
            txtSellIncl.Tag = "1"
            chkStock.Value = rs.Fields("Stock_Item")
            txtLink.Text = "<Not Linked>"
            If txtPackSize.Text = 1 Then
                cmdLink.Enabled = False
                fmLink.ForeColor = &HE0E0E0
                txtLink.Enabled = False
                picLink.BorderColor = &HE0E0E0
            Else
                If chkStock.Value = 0 Then
                    cmdLink.Enabled = True
                    fmLink.ForeColor = vbBlack
                    txtLink.Enabled = True
                    picLink.BorderColor = vbBlack
                    ActiveReadServer2 "Select * from Pack_Links where Product_Code ='" & txtProductCode.Text & "'"
                    If rs2.RecordCount > 0 Then
                        txtLink.Text = rs2.Fields("Link_Code") & ""
                    Else
                        txtLink.Text = "<Not Linked>"
                    End If
                Else
                    cmdLink.Enabled = False
                    fmLink.ForeColor = &HE0E0E0
                    txtLink.Enabled = False
                    picLink.BorderColor = &HE0E0E0
                End If
            End If
            If Abs(chkStock.Value) = 1 Then
                cmdLinks.Visible = True
            Else
                cmdLinks.Visible = False
            End If
            chkRecipe.Value = rs.Fields("Recipe_Item")
            If Abs(chkRecipe.Value) = 1 Then
                cmdPrep.Visible = True
                cmdPrint.Visible = True
            Else
                cmdPrep.Visible = False
                cmdPrint.Visible = fasle
            End If
            chkSales.Value = rs.Fields("Sales_Item")
            
            chkTouch.Value = rs.Fields("Touch_Item")
            chkProduction.Value = Val(rs.Fields("Production_Item") & "")
            txtEmpty.Text = Val(rs.Fields("Weight_Empty") & "")
            txtFull.Text = Val(rs.Fields("Weight_Full") & "")
            chkDeposit.Value = rs.Fields("Returnable_Item")
            chkScale.Value = rs.Fields("Scale_Item")
            cmbembedtype.Value = rs.Fields("Scaleitemtype")
            
            
            chkDelete.Value = rs.Fields("Once_Off")
            chkWhole.Value = Val(rs.Fields("Whole_Unit") & "")
            If chkSales.Value = 1 Then
                txtSellExcl.Tag = "1"
                Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
                If txtSellIncl.Text <> "N/A" Then
                    txtSellExcl.Text = Format(txtSellIncl.Text / ((100 + Tax) / 100), "0.00")
                End If
                txtSellExcl.Tag = ""
                If Val(txtLandCost.Text) <> 0 Then
                    txtMarkup.Tag = "1"
                    txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                    txtMarkup.Tag = ""
                End If
                If Val(txtSellExcl) <> 0 Then
                    txtGross.Tag = "1"
                    txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
                    txtGross.Tag = ""
                Else
                    txtGross.Tag = "1"
                    txtGross.Text = "0"
                    txtGross.Tag = ""
                End If
                chkTouch.Enabled = True
                txtSellExcl.Enabled = True
                txtSellIncl.Enabled = True
                txtMarkup.Enabled = True
                txtGross.Enabled = True
                cmbTax.Enabled = True
                ButtonEx3.Enabled = True
            End If
            txtSellIncl.Tag = ""
        End If
        rs.Close
        grdPrices.TextMatrix(7, 1) = "0.00"
        grdPrices.TextMatrix(7, 2) = "0.00"
        grdPrices.TextMatrix(7, 3) = "0.00"
        grdPrices.TextMatrix(7, 4) = "0.00"
        grdPrices.TextMatrix(7, 5) = "0.00"
        
        grdPrices.TextMatrix(3, 1) = "0"
        grdPrices.TextMatrix(3, 2) = "0"
        grdPrices.TextMatrix(3, 3) = "0"
        grdPrices.TextMatrix(3, 4) = "0"
        grdPrices.TextMatrix(3, 5) = "0"
        
        grdPrices.TextMatrix(4, 1) = "0"
        grdPrices.TextMatrix(4, 2) = "0"
        grdPrices.TextMatrix(4, 3) = "0"
        grdPrices.TextMatrix(4, 4) = "0"
        grdPrices.TextMatrix(4, 5) = "0"
        cmdTestR.Visible = False
        For i = 0 To cmdTab.Count - 1
            picBox(i).Visible = False
            cmdTab(i).Value = 0
        Next i
        cmdTab(0).Value = 1
        picBox(0).Visible = True
        
        ActiveReadServer "select * from Product_Prices where Product_Code= '" & txtProductCode.Text & "'"
        If rs.RecordCount > 0 Then
            grdPrices.TextMatrix(7, 1) = Format(rs.Fields("Price2"), "0.00")
            grdPrices.TextMatrix(7, 2) = Format(rs.Fields("Price3"), "0.00")
            grdPrices.TextMatrix(7, 3) = Format(rs.Fields("Price4"), "0.00")
            grdPrices.TextMatrix(7, 4) = Format(rs.Fields("Price5"), "0.00")
            grdPrices.TextMatrix(7, 5) = Format(rs.Fields("Price6"), "0.00")
        End If
        rs.Close
    End If
    grdRecipe.Rows = 1
    Select Case chkRecipe.Value
        Case 0
            cmdTab(1).Enabled = False
            txtLandCost.Locked = False
            txtLandCost.Enabled = True
            lblCost.Caption = "Landed Cost:"
        Case 1
            cmdTab(1).Enabled = True
            Select Case cmbUnit.Text
                Case "Preparation Recipe"
                    txtLandCost.Enabled = False
                    txtLandCost.Text = "N/A"
                Case Else
                    txtLandCost.Locked = True
                    lblCost.Caption = "Theoretical Cost:"
            End Select
            
            ActiveReadServer "Select * from Recipes where Product_Code= '" & txtProductCode.Text & "' order by Line_No"
            While Not rs.EOF
                grdRecipe.Rows = grdRecipe.Rows + 1
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1) = rs.Fields("Description")
                Select Case rs.Fields("Line_Type")
                    Case 0
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Message"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                        grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                        grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
                    Case 1
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Preparation Recipe"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                        grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                        grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
                    Case 2
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Sales Item"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 3
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 4
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item (Hidden)"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 5
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Price/Size Change"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = "0.00"
                        grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                        grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
                    Case 6
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Sales Item (Choice)"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 7
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item (Choice)"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 8
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Exit"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                        grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                        grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
                End Select
                rs.MoveNext
            Wend
            rs.Close
    End Select
    If grdRecipe.Rows = 1 Then grdRecipe.Rows = 2
    On Error GoTo 0
End Sub

Private Sub grdProd_AfterSort(ByVal Col As Long, Order As Integer)
    On Error Resume Next
    ActiveReadServer "Select * from Products where Product_Code='" & grdProd.TextMatrix(grdProd.Row, 0) & "'"
    If rs.RecordCount > 0 Then
        txtProductCode.Text = rs.Fields("Product_Code")
        TopRow = 0
        BottomRow = 0
        For i = 1 To Len(txtProductCode.Text) - 1
            If Len(txtProductCode.Text) < 9 Then
                If i / 2 <> Int(i / 2) Then
                    TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
                Else
                    BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
                End If
            Else
                If i / 2 = Int(i / 2) Then
                    TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
                Else
                    BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
                End If
            End If
        Next i
        TopRow = TopRow * 3
        Result = TopRow + BottomRow
        Result = (1 - ((Result / 10) - Int((Result / 10)))) * 10
        If Result = 10 Then Result = 0
        If Round(Result, 0) = Int(Val(Right(txtProductCode, 1))) Then
            PicBC.Visible = True
        Else
            PicBC.Visible = False
        End If
        txtDescription.Text = rs.Fields("Description")
        txtPackSize.Text = rs.Fields("Pack_Size")
        txtShort.Text = rs.Fields("Description")
        txtRef.Text = rs.Fields("Nappi_Code") & ""
        If rs.Fields("Unit_of_Measure") & "" = "" Then
            cmbUnit.Text = "each"
        Else
            cmbUnit.Text = rs.Fields("Unit_of_Measure")
        End If
        If rs.Fields("Unit_Size") = 0 Then
            txtUnitSize.Text = ""
        Else
            txtUnitSize.Text = rs.Fields("Unit_Size") & ""
        End If
        If grdProd.TextMatrix(grdProd.Row, 2) = "0" Or grdProd.TextMatrix(grdProd.Row, 2) = "" Then
            cmbDepartments.Text = "<Unbound>"
        Else
            cmbDepartments.Text = grdProd.TextMatrix(grdProd.Row, 2)
        End If
        txtLandCost.Text = grdProd.TextMatrix(grdProd.Row, 4)
        txtLandCost.ToolTipText = " Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & " "
        cmbTax.Text = grdProd.TextMatrix(grdProd.Row, 5)
        txtSellIncl.Text = Format(grdProd.TextMatrix(grdProd.Row, 6), "0.00")
        If rs.Fields("Scale_Prefix") & "" = "" Then
            cmbScalePrefix.Text = "20"
        Else
            cmbScalePrefix.Text = rs.Fields("Scale_Prefix") & ""
        End If
        If rs.Fields("Kitchen1") & "" = "" Then
            cmbPrinter1.Text = "<None>"
        Else
            cmbPrinter1.Text = rs.Fields("Kitchen1")
        End If
        If rs.Fields("Kitchen2") & "" = "" Then
            cmbPrinter2.Text = "<None>"
        Else
            cmbPrinter2.Text = rs.Fields("Kitchen2")
        End If
        chkStock.Value = rs.Fields("Stock_Item")
        If txtPackSize.Text = 1 Then
            cmdLink.Enabled = False
            fmLink.ForeColor = &HE0E0E0
            txtLink.Enabled = False
            picLink.BorderColor = &HE0E0E0
        Else
            If chkStock.Value = 0 Then
                cmdLink.Enabled = True
                fmLink.ForeColor = vbBlack
                txtLink.Enabled = True
                picLink.BorderColor = vbBlack
            Else
                cmdLink.Enabled = False
                fmLink.ForeColor = &HE0E0E0
                txtLink.Enabled = False
                picLink.BorderColor = &HE0E0E0
            End If
        End If
        If Abs(chkStock.Value) = 1 Then
            cmdLinks.Visible = True
        Else
            cmdLinks.Visible = False
        End If
        chkSales.Value = rs.Fields("Sales_Item")
        chkRecipe.Value = rs.Fields("Recipe_Item")
        If Abs(chkRecipe.Value) = 1 Then
            cmdPrep.Visible = True
            cmdPrint.Visible = True
        Else
            cmdPrep.Visible = False
            cmdPrint.Visible = False
        End If
        chkTouch.Value = rs.Fields("Touch_Item")
        chkDeposit.Value = rs.Fields("Returnable_Item")
        chkProduction.Value = Val(rs.Fields("Production_Item") & "")
        chkScale.Value = rs.Fields("Scale_Item")
        chkDelete.Value = rs.Fields("Once_Off")
        chkWhole.Value = Val(rs.Fields("Whole_Unit") & "")
        txtEmpty.Text = Val(rs.Fields("Weight_Empty") & "")
        txtFull.Text = Val(rs.Fields("Weight_Full") & "")
    End If
    rs.Close
    grdRecipe.Rows = 1
    Select Case chkRecipe.Value
        Case 0
            cmdTab(1).Enabled = False
            txtLandCost.Locked = False
            txtLandCost.Enabled = True
            lblCost.Caption = "Landed Cost:"
        Case 1
            cmdTab(1).Enabled = True
            Select Case cmbUnit.Text
                Case "Preparation Recipe"
                    txtLandCost.Enabled = False
                    txtLandCost.Text = "N/A"
                Case Else
                    txtLandCost.Locked = True
                    lblCost.Caption = "Theoretical Cost:"
            End Select
            ActiveReadServer "Select * from Recipes where Product_Code= '" & txtProductCode.Text & "' order by Line_No"
            While Not rs.EOF
                grdRecipe.Rows = grdRecipe.Rows + 1
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1) = rs.Fields("Description")
                Select Case rs.Fields("Line_Type")
                    Case 0
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Message"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                        grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                        grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
                    Case 1
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Preparation Recipe"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                        grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                        grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
                    Case 2
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Sales Item"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 3
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 4
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item (Hidden)"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 5
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Price/Size Change"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = "0.00"
                        grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                        grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
                    Case 6
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Sales Item (Choice)"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 7
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item (Choice)"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                        ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                        If rs1.RecordCount > 0 Then
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                        End If
                        rs1.Close
                    Case 8
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Exit"
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                        grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                        grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
                End Select
                rs.MoveNext
            Wend
            rs.Close
    End Select
    If grdRecipe.Rows = 1 Then grdRecipe.Rows = 2
    On Error GoTo 0
End Sub
Private Sub grdProd_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub
Private Sub grdRecipe_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If grdRecipe.Tag = grdRecipe.Text Then Exit Sub
    Select Case Col
        Case 0
            Select Case grdRecipe.TextMatrix(Row, 0)
                Case "Message", "Preparation Recipe", "Price/Size Change", "Exit"
                    grdRecipe.TextMatrix(Row, 1) = ""
                    grdRecipe.TextMatrix(Row, 2) = " "
                    grdRecipe.TextMatrix(Row, 3) = " "
                    grdRecipe.TextMatrix(Row, 4) = "0.00"
                    grdRecipe.MergeRow(Row) = True
                    grdRecipe.Cell(flexcpBackColor, Row, 2, Row, 4) = &HE0E0E0
                Case Else
                    grdRecipe.TextMatrix(Row, 1) = ""
                    grdRecipe.TextMatrix(Row, 2) = ""
                    grdRecipe.TextMatrix(Row, 3) = ""
                    grdRecipe.TextMatrix(Row, 4) = ""
                    grdRecipe.MergeRow(Row) = False
                    grdRecipe.Cell(flexcpBackColor, Row, 2, Row, 4) = grdRecipe.BackColor
            End Select
        Case 1
            Select Case grdRecipe.TextMatrix(Row, 0)
                Case "Message", "Price/Size Change", "Exit"
                    grdRecipe.TextMatrix(Row, 1) = UCase(Left(grdRecipe.TextMatrix(Row, 1), 1)) & Mid(grdRecipe.TextMatrix(Row, 1), 2)
                Case "Preparation Recipe"
                Case "Sales Item", "Stock Item", "Stock Item (Hidden)", "Stock Item (Choice)", "Sales Item (Choice)"
                    If InStrRev(grdRecipe.TextMatrix(Row, 1), ",") = 0 Then
                        grdRecipe.TextMatrix(Row, 1) = ""
                    Else
                        ActiveReadServer "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(Row, 1), InStrRev(grdRecipe.TextMatrix(Row, 1), ",") + 1) & "'"
                        If rs.RecordCount > 0 Then
                            Select Case rs.Fields("Unit_of_Measure")
                                Case "ml"
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.ColComboList(2) = "ml|Single Tot|Double Tot"
                                    grdRecipe.TextMatrix(Row, 2) = "ml"
                                    grdRecipe.TextMatrix(Row, 3) = rs.Fields("Unit_Size")
                                Case "lt"
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.ColComboList(2) = "ml|lt|Single Tot|Double Tot"
                                    grdRecipe.TextMatrix(Row, 2) = "ml"
                                    grdRecipe.TextMatrix(Row, 3) = rs.Fields("Unit_Size") * 1000
                                Case "g"
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.TextMatrix(Row, 2) = "g"
                                    grdRecipe.TextMatrix(Row, 3) = rs.Fields("Unit_Size")
                                Case "kg"
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.ColComboList(2) = "g|kg"
                                    grdRecipe.TextMatrix(Row, 2) = "g"
                                    grdRecipe.TextMatrix(Row, 3) = rs.Fields("Unit_Size") * 1000
                                Case Else
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.TextMatrix(Row, 2) = "each"
                                    grdRecipe.TextMatrix(Row, 3) = "1"
                            End Select
                        End If
                        grdRecipe.TextMatrix(Row, 5) = rs.Fields("Unit_Size") & ""
                        If grdRecipe.TextMatrix(Row, 5) = 0 Then grdRecipe.TextMatrix(Row, 5) = "1"
                        grdRecipe.TextMatrix(Row, 6) = rs.Fields("Unit_of_Measure") & ""
                        grdRecipe.TextMatrix(Row, 4) = 1
                        grdRecipe.TextMatrix(Row, 4) = Format(rs.Fields("Ave_Cost"), "0.00")
                        grdRecipe.TextMatrix(Row, 7) = Format(rs.Fields("Ave_Cost"), "0.00")
                        rs.Close
                    End If
            End Select
        Case 2
            If grdRecipe.Tag = grdRecipe.TextMatrix(Row, 2) Then Exit Sub
            Select Case grdRecipe.Tag
                Case "ml"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "lt"
                            grdRecipe.TextMatrix(Row, 3) = Val(grdRecipe.TextMatrix(Row, 3)) / 1000
                        Case "1 x 25ml"
                            grdRecipe.TextMatrix(Row, 3) = "25"
                         Case "2 x 50ml"
                            grdRecipe.TextMatrix(Row, 3) = "50"
                        Case "Whole Unit"
                            grdRecipe.TextMatrix(Row, 3) = grdRecipe.TextMatrix(Row, 5)
                    End Select
                Case "lt"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "ml"
                            grdRecipe.TextMatrix(Row, 3) = Val(grdRecipe.TextMatrix(Row, 3)) * 1000
                        Case "1 x 25ml"
                            grdRecipe.TextMatrix(Row, 3) = "0.025"
                         Case "2 x 50ml"
                            grdRecipe.TextMatrix(Row, 3) = "0.050"
                        Case "Whole Unit"
                            grdRecipe.TextMatrix(Row, 3) = grdRecipe.TextMatrix(Row, 5)
                    End Select
                Case "g"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "kg"
                            grdRecipe.TextMatrix(Row, 3) = Val(grdRecipe.TextMatrix(Row, 3)) / 1000
                    End Select
                Case "kg"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "g"
                            grdRecipe.TextMatrix(Row, 3) = Val(grdRecipe.TextMatrix(Row, 3)) * 1000
                    End Select
                Case "Single Tot"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "ml"
                            grdRecipe.TextMatrix(Row, 3) = "25"
                        Case "lt"
                            grdRecipe.TextMatrix(Row, 3) = "0.025"
                        Case "Whole Unit"
                            grdRecipe.TextMatrix(Row, 3) = grdRecipe.TextMatrix(Row, 5)
                    End Select
                Case "Double Tot"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "ml"
                            grdRecipe.TextMatrix(Row, 3) = "50"
                        Case "lt"
                            grdRecipe.TextMatrix(Row, 3) = "0.050"
                        Case "Whole Unit"
                            grdRecipe.TextMatrix(Row, 3) = grdRecipe.TextMatrix(Row, 5)
                    End Select
            End Select
            If grdRecipe.TextMatrix(Row, 2) = "Single Tot" Then grdRecipe.TextMatrix(Row, 3) = "1 x 25ml"
            If grdRecipe.TextMatrix(Row, 2) = "Double Tot" Then grdRecipe.TextMatrix(Row, 3) = "2 x 25ml"
            If grdRecipe.TextMatrix(Row, 2) = "Whole Unit" Then grdRecipe.TextMatrix(Row, 3) = 1
        Case 3
            ActiveReadServer "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(Row, 1), InStrRev(grdRecipe.TextMatrix(Row, 1), ",") + 1) & "'"
            If rs.RecordCount > 0 Then
                grdRecipe.TextMatrix(Row, 5) = rs.Fields("Unit_Size") & ""
                If grdRecipe.TextMatrix(Row, 5) = 0 Then grdRecipe.TextMatrix(Row, 5) = "1"
                grdRecipe.TextMatrix(Row, 6) = rs.Fields("Unit_of_Measure") & ""
                grdRecipe.TextMatrix(Row, 4) = 1
                grdRecipe.TextMatrix(Row, 4) = Format(rs.Fields("Ave_Cost"), "0.00")
                grdRecipe.TextMatrix(Row, 7) = Format(rs.Fields("Ave_Cost"), "0.00")
            End If
            rs.Close
        Case 4
            grdRecipe.TextMatrix(Row, 4) = Format(grdRecipe.TextMatrix(Row, 4), "0.00")
    End Select
    If Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) > 0 Then
        Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
            Case "Stock Item", "Sales Item", "Stock Item (Hidden)", "Stock Item (Choice)", "Sales Item (Choice)"
                Select Case grdRecipe.TextMatrix(Row, 6)
                    Case "ml"
                        Select Case grdRecipe.TextMatrix(Row, 2)
                            Case "ml"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "Single Tot"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (25 / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "Double Tot"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (50 / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "Whole Unit"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)), "0.000")
                        End Select
                    Case "lt"
                        Select Case grdRecipe.TextMatrix(Row, 2)
                            Case "lt"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "ml"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * ((Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5)) / 1000)), "0.000")
                            Case "Single Tot"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * ((25 / Val(grdRecipe.TextMatrix(Row, 5)) / 1000)), "0.000")
                            Case "Double Tot"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * ((50 / Val(grdRecipe.TextMatrix(Row, 5)) / 1000)), "0.000")
                            Case "Whole Unit"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)), "0.000")
                        End Select
                    Case "g"
                        Select Case grdRecipe.TextMatrix(Row, 2)
                            Case "g"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                        End Select
                    Case "kg"
                        Select Case grdRecipe.TextMatrix(Row, 2)
                            Case "kg"
                                If Val(grdRecipe.TextMatrix(Row, 5)) <> 0 Then
                                    grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                                Else
                                    grdRecipe.TextMatrix(Row, 4) = "0.00"
                                End If
                            Case "g"
                                If Val(grdRecipe.TextMatrix(Row, 5)) <> 0 Then
                                    grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * ((Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5)) / 1000)), "0.000")
                                Else
                                    grdRecipe.TextMatrix(Row, 4) = "0.00"
                                End If
                        End Select
                    Case "each"
                        grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3))), "0.000")
                End Select
        End Select
    End If
    totalcost = 0
    For i = 1 To grdRecipe.Rows - 1
        Select Case grdRecipe.TextMatrix(i, 0)
            Case "Stock Item", "Sales Item", "Stock Item (Hidden)"
                totalcost = totalcost + Val(grdRecipe.TextMatrix(i, 4))
        End Select
    Next i
    If totalcost > 0 Then
        txtLandCost.Text = Format(totalcost, "0.00")
    End If
End Sub
Private Sub grdRecipe_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdRecipe.Tag = grdRecipe.Text
End Sub

Private Sub grdRecipe_Click()
    If grdRecipe.Col = 4 Then
        totalcost = 0
        For i = 1 To grdRecipe.Rows - 1
            totalcost = totalcost + Val(grdRecipe.TextMatrix(grdRecipe.Row, 4))
        Next i
        If totalcost > 0 And grdRecipe.Col = 4 Then
            Timer1.Enabled = True
        End If
    End If
End Sub

Private Sub grdRecipe_EnterCell()
    totalcost = 0
    For i = 1 To grdRecipe.Rows - 1
        totalcost = totalcost + Val(grdRecipe.TextMatrix(i, 4))
    Next i
    If totalcost > 0 And grdRecipe.Col = 4 Then
        Timer1.Enabled = True
    End If
    Select Case grdRecipe.Col
        Case 0
            grdRecipe.ColComboList(0) = "Message|Preparation Recipe|Sales Item|Sales Item (Choice)|Stock Item|Stock Item (Choice)|Stock Item (Hidden)|Price/Size Change|Exit"
            If grdRecipe.Text = "" Then
                grdRecipe.Text = "Message"
                grdRecipe.TextMatrix(grdRecipe.Row, 2) = " "
                grdRecipe.TextMatrix(grdRecipe.Row, 3) = " "
                grdRecipe.TextMatrix(grdRecipe.Row, 4) = " "
                grdRecipe.MergeRow(grdRecipe.Row) = True
                grdRecipe.Cell(flexcpBackColor, grdRecipe.Row, 2, grdRecipe.Row, 4) = &HE0E0E0
            End If
            grdRecipe.Editable = flexEDKbdMouse
            grdRecipe.ColComboList(1) = ""
        Case 1
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Message", "Price Size Change", "Exit"
                    grdRecipe.ColComboList(1) = ""
                    grdRecipe.Editable = flexEDKbdMouse
                    grdRecipe.ColComboList(2) = ""
                Case "Preparation Recipe"
                    If grdRecipe.ColComboList(1) = "" Then
                        Screen.MousePointer = 11
                        ActiveReadServer "Select * from Products where Unit_of_Measure ='Preparation Recipe' order by Description"
                        grdRecipe.ColComboList(1) = ""
                        While Not rs.EOF
                            grdRecipe.ColComboList(1) = grdRecipe.ColComboList(1) & "|" & rs.Fields("Description") & " ," & rs.Fields("Product_Code")
                            rs.MoveNext
                        Wend
                        rs.Close
                        Screen.MousePointer = 0
                    End If
                    grdRecipe.Editable = flexEDKbdMouse
                    grdRecipe.ColComboList(2) = ""
                Case "Stock Item", "Stock Item (Choice)"
                    If grdRecipe.ColComboList(1) = "" Then
                        Screen.MousePointer = 11
                        ActiveReadServer "Select Product_Code,Description,isnull(Unit_Size,'') as Unit_Size,Unit_of_Measure,Ave_Cost from Products where Stock_Item=1 order by Description"
                        While Not rs.EOF
                            If rs.Fields("Unit_Size") = "0" Then
                                UnitSize = ""
                            Else
                                UnitSize = rs.Fields("Unit_Size")
                            End If
                            grdRecipe.ColComboList(1) = grdRecipe.ColComboList(1) & "|" & rs.Fields("Description") & " " & UnitSize & rs.Fields("Unit_of_Measure") & " ," & rs.Fields("Product_Code")
                            rs.MoveNext
                        Wend
                        rs.Close
                        Screen.MousePointer = 0
                    End If
                    grdRecipe.Editable = flexEDKbdMouse
                Case "Stock Item (Hidden)"
                    If grdRecipe.ColComboList(1) = "" Then
                        Screen.MousePointer = 11
                        ActiveReadServer "Select Product_Code,Description,isnull(Unit_Size,'') as Unit_Size,Unit_of_Measure,Ave_Cost from Products where Stock_Item=1 order by Description"
                        While Not rs.EOF
                            If rs.Fields("Unit_Size") = "0" Then
                                UnitSize = ""
                            Else
                                UnitSize = rs.Fields("Unit_Size")
                            End If
                            grdRecipe.ColComboList(1) = grdRecipe.ColComboList(1) & "|" & rs.Fields("Description") & " " & UnitSize & rs.Fields("Unit_of_Measure") & " ," & rs.Fields("Product_Code")
                            rs.MoveNext
                        Wend
                        rs.Close
                        Screen.MousePointer = 0
                    End If
                    grdRecipe.Editable = flexEDKbdMouse
                Case "Sales Item", "Sales Item (Choice)"
                    If grdRecipe.ColComboList(1) = "" Then
                        Screen.MousePointer = 11
                        ActiveReadServer "Select Product_Code,Description,isnull(Unit_Size,'') as Unit_Size,Unit_of_Measure,Ave_Cost from Products where Sales_Item=1 order by Description"
                        While Not rs.EOF
                            If rs.Fields("Unit_Size") = "0" Then
                                UnitSize = ""
                            Else
                                UnitSize = rs.Fields("Unit_Size")
                            End If
                            grdRecipe.ColComboList(1) = grdRecipe.ColComboList(1) & "|" & rs.Fields("Description") & " " & UnitSize & rs.Fields("Unit_of_Measure") & " ," & rs.Fields("Product_Code")
                            rs.MoveNext
                        Wend
                        rs.Close
                        Screen.MousePointer = 0
                    End If
                    grdRecipe.Editable = flexEDKbdMouse
            End Select
        Case 2
            If grdRecipe.TextMatrix(grdRecipe.Row, 1) = "" Then
                grdRecipe.Col = 1
                Exit Sub
            End If
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Stock Item", "Sales Item", "Stock Item (Hidden)", "Stock Item (Choice)", "Sales Item (Choice)"
                    ActiveReadServer "Select Unit_of_Measure from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Row, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Row, 1), ",") + 1) & "'"
                    If rs.RecordCount > 0 Then
                        Select Case rs.Fields("Unit_of_Measure")
                            Case "ml"
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.ColComboList(2) = "ml|Single Tot|Double Tot|Whole Unit"
                                grdRecipe.Editable = flexEDKbdMouse
                            Case "lt"
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.ColComboList(2) = "ml|lt|Single Tot|Double Tot|Whole Unit"
                                grdRecipe.Editable = flexEDKbdMouse
                            Case "g"
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.ColComboList(2) = "g|Whole Unit"
                                grdRecipe.Editable = flexEDKbdMouse
                            Case "kg"
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.ColComboList(2) = "g|kg|Whole Unit"
                                grdRecipe.Editable = flexEDKbdMouse
                            Case Else
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.Editable = flexEDNone
                        End Select
                    End If
                    rs.Close
                Case Else
                    grdRecipe.Editable = flexEDNone
            End Select
        Case 3
            If grdRecipe.TextMatrix(grdRecipe.Row, 1) = "" Then
                grdRecipe.Col = 1
                Exit Sub
            End If
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Message", "Preparation Recipe", "Price/Size Change", "Exit"
                    grdRecipe.Editable = flexEDNone
                    grdRecipe.ColComboList(2) = ""
                Case Else
                    grdRecipe.Editable = flexEDKbdMouse
            End Select
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 2)
                Case "Whole Unit", "Single Tot", "Double Tot"
                    grdRecipe.Editable = flexEDNone
            End Select
        Case 4
            If grdRecipe.TextMatrix(grdRecipe.Row, 1) = "" Then
                grdRecipe.Col = 1
                Exit Sub
            End If
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Message", "Preparation Recipe", "Price/Size Change", "Exit"
                    grdRecipe.Editable = flexEDNone
                    grdRecipe.ColComboList(2) = ""
                Case "Sales Item (Choice)"
                    grdRecipe.Editable = flexEDKbdMouse
                Case Else
                    grdRecipe.Editable = flexEDNone
            End Select
    End Select
End Sub

Private Sub grdRecipe_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46
            If grdRecipe.Row <> 0 Then
                grdRecipe.RemoveItem grdRecipe.Row
                If grdRecipe.Rows = 1 Then
                    grdRecipe.Rows = grdRecipe.Rows + 1
                End If
                totalcost = 0
                For i = 1 To grdRecipe.Rows - 1
                    totalcost = totalcost + Val(grdRecipe.TextMatrix(i, 4))
                Next i
                If totalcost > 0 Then
                    txtLandCost.Text = Format(totalcost, "0.00")
                End If
            End If
        Case 40
            If grdRecipe.Row = grdRecipe.Rows - 1 And grdRecipe.TextMatrix(grdRecipe.Row, 1) <> "" Then
                If grdRecipe.Rows < 46 Then
                    grdRecipe.Rows = grdRecipe.Rows + 1
                    grdRecipe.Col = 0
                End If
            Else
                If grdRecipe.TextMatrix(grdRecipe.Row, 0) = "Exit" Then
                    If grdRecipe.Rows < 46 Then
                        grdRecipe.Rows = grdRecipe.Rows + 1
                        grdRecipe.Col = 0
                    End If
                End If
            End If
        Case 38
            If grdRecipe.TextMatrix(grdRecipe.Row, 1) = "" And grdRecipe.Row = grdRecipe.Rows - 1 Then
                grdRecipe.RemoveItem grdRecipe.Row
            End If
    End Select
End Sub

Private Sub grdRecipe_KeyPress(KeyAscii As Integer)
    Select Case grdRecipe.Col
        Case 0
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
        Case 1
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Stock Item", "Sales Item", "Preparation Recipe", "Stock Item (Hidden)"
                    Select Case KeyAscii
                    Case 8, 13, 27
                    Case Else
                        KeyAscii = 0
                End Select
            End Select
        Case 2
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
        Case 3
            If grdRecipe.TextMatrix(grdRecipe.Row, 0) = "Stock Item" Or grdRecipe.TextMatrix(grdRecipe.Row, 0) = "Stock Item (Hidden)" Then
                Select Case KeyAscii
                    Case 8, 13, 27
                    Case 46, 48 To 57
                    Case Else
                        KeyAscii = 0
                End Select
            Else
                Select Case KeyAscii
                    Case 8, 13, 27
                    Case 48 To 57
                    Case Else
                        KeyAscii = 0
                End Select
            End If
        Case 4
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
    End Select
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub grdRecipe_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case grdRecipe.Col
        Case 0
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
        Case 2
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
        Case 3
            If grdRecipe.TextMatrix(Row, 0) = "Stock Item" Or grdRecipe.TextMatrix(Row, 0) = "Stock Item (Hidden)" Then
                Select Case KeyAscii
                    Case 8, 13, 27
                    Case 46, 48 To 57
                    Case Else
                        KeyAscii = 0
                End Select
            Else
                Select Case KeyAscii
                    Case 8, 13, 27
                    Case 48 To 57
                    Case Else
                        KeyAscii = 0
                End Select
            End If
            DoEvents
            If InStr(grdRecipe.EditText, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
        Case 4
            Select Case KeyAscii
                Case 48 To 57
                Case 8, 13, 27, 45, 46
                Case Else
                    KeyAscii = 0
            End Select
            If InStr(grdRecipe.EditText, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    End Select
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub grdRecipe_LostFocus()
    Timer1.Enabled = False
    txtLandCost.BackColor = &H80000005
End Sub
Private Sub grdRecipe_RowColChange()
    If grdRecipe.Col <> 4 Then
        Timer1.Enabled = False
        txtLandCost.BackColor = &H80000005
    End If
End Sub

Private Sub grdSub1_Click()
    If grdSub1.Rows > 1 Then
        For i = 1 To grdMinor1.Rows - 1
            grdMinor1.TextMatrix(i, 2) = "0"
            grdMinor1.Cell(flexcpFontItalic, i, 0, i, 1) = False
            grdMinor1.Cell(flexcpFontBold, i, 0, i, 1) = False
            grdMinor1.Cell(flexcpBackColor, i, 0, i, 2) = ""
        Next i
        
        ActiveReadServer1 "Select * from Department_links where Dept_No = '" & grdSub1.TextMatrix(grdSub1.Row, 0) & "'"
        While Not rs1.EOF
            For i = 1 To grdMinor1.Rows - 1
                If rs1.Fields("Location_No") = grdMinor1.TextMatrix(i, 0) Then
                    grdMinor1.TextMatrix(i, 2) = "1"
                    grdMinor1.Cell(flexcpFontItalic, i, 0, i, 1) = True
                    grdMinor1.Cell(flexcpFontBold, i, 0, i, 1) = True
                    grdMinor1.Cell(flexcpBackColor, i, 0, i, 2) = &HFFC0C0
                End If
            Next i
            rs1.MoveNext
        Wend
        rs1.Close
        If grdSub1.Row <> 0 Then
            cmbDepartments.Text = grdSub1.TextMatrix(grdSub1.Row, 0) & " - " & grdSub1.TextMatrix(grdSub1.Row, 1)
        End If
    End If
End Sub
Private Sub grdSub1_RowColChange()
    grdSub1_Click
End Sub

Private Sub ButtonEx2_Click()
    On Error Resume Next
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
    frmLabel.Show vbModal
End Sub
Private Sub Timer1_Timer()
    Select Case txtLandCost.BackColor
        Case &HD9FFFC
            txtLandCost.BackColor = &H80000005
        Case &H80000005
            txtLandCost.BackColor = &HD9FFFC
    End Select
End Sub
Private Sub ButtonEx3_Click()
    On Error Resume Next
    Select Case picPrice.Visible
        Case True: picPrice.Visible = False
        Case False
            picPrice.Visible = True
            grdPrices.SetFocus
    End Select
    grdPrices.Row = 2
    grdPrices.Col = 1
    For i = 1 To 5
        grdPrices.TextMatrix(2, i) = txtLandCost.Text
        grdPrices.TextMatrix(6, i) = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
        grdPrices.TextMatrix(5, i) = Format(Val(grdPrices.TextMatrix(7, i)) / ((100 + Val(grdPrices.TextMatrix(6, i))) / 100), "0.00")
        If Val(grdPrices.TextMatrix(2, i)) <> 0 Then
            grdPrices.TextMatrix(3, i) = Round(((Val(grdPrices.TextMatrix(5, i)) - Val(grdPrices.TextMatrix(2, i))) / Val(grdPrices.TextMatrix(2, i)) * 100), 3)
        End If
        If Val(grdPrices.TextMatrix(5, i)) <> 0 Then
            grdPrices.TextMatrix(4, i) = Round(((Val(grdPrices.TextMatrix(5, i)) - Val(grdPrices.TextMatrix(2, i))) / Val(grdPrices.TextMatrix(5, i)) * 100), 3)
        End If
    Next i
    On Error GoTo 0
End Sub
Private Sub ButtonEx3_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
End Sub

Private Sub ButtonEx4_Click()
    Load_Products
End Sub
Private Sub txtDescription_GotFocus()
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub
Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtProductCode.SetFocus
        Case 40
            txtShort.SetFocus
    End Select
End Sub
Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 33 To 47, 58 To 64, 91 To 96, 123 To 127, 162 To 184, 247, 248, 191
            KeyAscii = 0
        
    End Select
End Sub

Private Sub txtDescription_LostFocus()
    On Error Resume Next
    CheckforSave
    txtDescription.Text = UCase(Left(txtDescription.Text, 1)) & Mid(txtDescription.Text, 2)
    If Len(txtDescription.Text) < 26 And txtShort.Text = "" Then
        txtShort.Text = txtDescription.Text
    End If
    On Error GoTo 0
End Sub

Private Sub txtGross_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub

Private Sub txtGross_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtMarkup.SetFocus
        Case 40
            txtSellExcl.SetFocus
    End Select
End Sub

Private Sub txtGross_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtGross_LostFocus()
    If txtGross.Text = "" Then txtGross.Text = "0"
End Sub
Private Sub txtLandCost_GotFocus()
    txtLandCost.SelStart = 0
    txtLandCost.SelLength = Len(txtLandCost.Text)
    If picDepartments.Visible = True Then picDepartments.Visible = False
    If picPrice.Visible = True Then picPrice.Visible = False
End Sub
Private Sub txtLandCost_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            cmbDepartments.SetFocus
        Case 40
            If txtMarkup.Enabled = True Then
                txtMarkup.SetFocus
            Else
                txtProductCode.SetFocus
            End If
    End Select
End Sub

Private Sub txtLandCost_KeyPress(KeyAscii As Integer)
    If InStr(txtLandCost.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtLandCost_LostFocus()
    If txtLandCost.Text = "" Then txtLandCost.Text = "0.00"
    txtLandCost.Text = Format(txtLandCost.Text, "0.00")
    Calculate_Plu "Landed Cost"
End Sub

