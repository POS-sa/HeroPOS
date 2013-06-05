VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDepartments 
   Caption         =   "Departments"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picProducts 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10275
      Index           =   1
      Left            =   0
      ScaleHeight     =   10275
      ScaleWidth      =   13875
      TabIndex        =   9
      Top             =   0
      Width           =   13875
      Begin VB.CheckBox chk_foreign_currency 
         BackColor       =   &H80000005&
         Caption         =   "Report on Forreign Currency"
         Height          =   255
         Left            =   9540
         TabIndex        =   34
         Top             =   1020
         Width           =   2355
      End
      Begin VB.PictureBox picFrame 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   4005
         TabIndex        =   14
         Top             =   3390
         Width           =   4005
         Begin MSForms.Image Image9 
            Height          =   90
            Left            =   0
            Top             =   360
            Width           =   3015
            BackColor       =   16761024
            Size            =   "5318;159"
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   8
            Left            =   3060
            Top             =   360
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   7
            Left            =   3390
            Top             =   360
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   6
            Left            =   3720
            Top             =   360
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin VB.Label Label7 
            Caption         =   "Department Tree."
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
            Left            =   90
            TabIndex        =   15
            Top             =   0
            Width           =   3915
         End
      End
      Begin VB.PictureBox picTree 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5325
         Left            =   0
         ScaleHeight     =   5325
         ScaleWidth      =   13755
         TabIndex        =   10
         Top             =   3900
         Width           =   13755
         Begin VSFlex8Ctl.VSFlexGrid grdSub 
            Height          =   4515
            Left            =   4785
            TabIndex        =   7
            Top             =   450
            Width           =   4665
            _cx             =   8229
            _cy             =   7964
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
            FormatString    =   $"frmDepartments.frx":0000
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
         Begin VSFlex8Ctl.VSFlexGrid grdMinor 
            Height          =   4515
            Left            =   9555
            TabIndex        =   8
            Top             =   450
            Width           =   4185
            _cx             =   7382
            _cy             =   7964
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmDepartments.frx":0078
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
            Editable        =   2
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
         Begin VSFlex8Ctl.VSFlexGrid grdMajor 
            Height          =   4515
            Left            =   0
            TabIndex        =   6
            Top             =   450
            Width           =   4680
            _cx             =   8255
            _cy             =   7964
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
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
            FormatString    =   $"frmDepartments.frx":00F0
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
         Begin VB.Label lblMaj 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Major Departments"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   120
            Width           =   4485
         End
         Begin VB.Label lblSub 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Departments"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4800
            TabIndex        =   12
            Top             =   120
            Width           =   4605
         End
         Begin VB.Label lblLinks 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Location Links"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   9570
            TabIndex        =   11
            Top             =   120
            Width           =   4725
         End
         Begin MSForms.Image Image11 
            Height          =   10005
            Left            =   4680
            Top             =   -30
            Width           =   135
            BorderStyle     =   0
            SpecialEffect   =   3
            Size            =   "238;17648"
         End
         Begin MSForms.Image Image12 
            Height          =   10005
            Left            =   9450
            Top             =   -60
            Width           =   135
            BorderStyle     =   0
            SpecialEffect   =   3
            Size            =   "238;17648"
         End
         Begin MSForms.Image picLoc 
            Height          =   495
            Left            =   9540
            Top             =   0
            Width           =   4215
            BorderStyle     =   0
            SpecialEffect   =   3
            Size            =   "7435;873"
         End
         Begin MSForms.Image picMin 
            Height          =   495
            Left            =   4770
            Top             =   0
            Width           =   4725
            BorderStyle     =   0
            SpecialEffect   =   3
            Size            =   "8334;873"
         End
         Begin MSForms.Image picMaj 
            Height          =   495
            Left            =   0
            Top             =   0
            Width           =   4725
            BackColor       =   14737632
            BorderStyle     =   0
            SpecialEffect   =   3
            Size            =   "8334;873"
         End
      End
      Begin btButtonEx.ButtonEx cmdUp2 
         Height          =   465
         Left            =   13080
         TabIndex        =   16
         Top             =   3390
         Width           =   585
         _ExtentX        =   1032
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
      Begin VB.Frame frmCodeRange 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Product Code Range"
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   5790
         TabIndex        =   17
         Top             =   1890
         Width           =   3075
         Begin VB.TextBox txtPluStop 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1740
            TabIndex        =   28
            Text            =   "0"
            Top             =   420
            Width           =   1155
         End
         Begin VB.TextBox txtPluStart 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   180
            TabIndex        =   27
            Text            =   "0"
            Top             =   420
            Width           =   1185
         End
         Begin MSForms.Image Image1 
            Height          =   345
            Index           =   1
            Left            =   1650
            Top             =   330
            Width           =   1305
            BackColor       =   16777215
            Size            =   "2302;609"
         End
         Begin MSForms.Image Image1 
            Height          =   345
            Index           =   0
            Left            =   90
            Top             =   330
            Width           =   1305
            BackColor       =   16777215
            Size            =   "2302;609"
         End
         Begin MSForms.Label Label2 
            Height          =   225
            Index           =   17
            Left            =   1380
            TabIndex        =   18
            Top             =   390
            Width           =   315
            BackColor       =   -2147483643
            Caption         =   "to"
            Size            =   "556;397"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fmOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Options"
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   5790
         TabIndex        =   29
         Top             =   1890
         Visible         =   0   'False
         Width           =   3075
         Begin VB.CheckBox chkZero 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Allow Zero Priced Sales Items"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   270
            TabIndex        =   30
            Top             =   270
            Width           =   2475
         End
      End
      Begin btButtonEx.ButtonEx cmdChange 
         Height          =   285
         Left            =   4110
         TabIndex        =   33
         Top             =   990
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         Appearance      =   3
         AutoMask        =   0   'False
         Enabled         =   0   'False
         Caption         =   "Change Department Number"
         Enabled         =   0   'False
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
      Begin VB.Label lblGL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GL Code: "
         Height          =   225
         Index           =   1
         Left            =   5940
         TabIndex        =   32
         Top             =   1050
         Visible         =   0   'False
         Width           =   1155
      End
      Begin MSForms.TextBox txtGL_Code 
         Height          =   285
         Left            =   7110
         TabIndex        =   31
         Top             =   990
         Visible         =   0   'False
         Width           =   1755
         VariousPropertyBits=   746604563
         MaxLength       =   30
         BorderStyle     =   1
         Size            =   "3096;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image picSeperate2 
         Height          =   135
         Left            =   -60
         Top             =   3210
         Width           =   13845
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "24421;238"
      End
      Begin MSForms.Image pic1 
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
         Index           =   27
         Left            =   900
         TabIndex        =   26
         Top             =   240
         Width           =   3105
         ForeColor       =   0
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Department Details"
         Size            =   "5477;450"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label6 
         Height          =   225
         Left            =   930
         TabIndex        =   25
         Top             =   1050
         Width           =   1725
         BackColor       =   -2147483643
         Caption         =   "Department Number:"
         Size            =   "3043;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   26
         Left            =   1260
         TabIndex        =   24
         Top             =   1395
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Department Name:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   25
         Left            =   1260
         TabIndex        =   23
         Top             =   1740
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Short Name:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   24
         Left            =   1260
         TabIndex        =   22
         Top             =   2085
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Type:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   1410
         X2              =   7590
         Y1              =   750
         Y2              =   750
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   23
         Left            =   90
         TabIndex        =   21
         Top             =   660
         Width           =   1785
         ForeColor       =   12582912
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "General"
         Size            =   "3149;450"
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
         Index           =   14
         Left            =   1260
         TabIndex        =   20
         Top             =   2400
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Parent Department:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblType 
         Height          =   285
         Left            =   2895
         TabIndex        =   3
         Top             =   2055
         Width           =   2730
         BackColor       =   -2147483643
         VariousPropertyBits=   8388627
         Caption         =   "Root Department"
         Size            =   "4815;503"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblParent 
         Height          =   165
         Left            =   2895
         TabIndex        =   4
         Top             =   2430
         Width           =   2730
         BackColor       =   -2147483643
         VariousPropertyBits=   8388627
         Caption         =   "None"
         Size            =   "4815;291"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image picTopBar2 
         Height          =   645
         Left            =   0
         Top             =   3300
         Width           =   13770
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "24289;1138"
      End
      Begin MSForms.ComboBox cmbTax1 
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Tag             =   "Up"
         Top             =   2715
         Width           =   2955
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "5212;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblTax1 
         Height          =   225
         Left            =   1260
         TabIndex        =   19
         Top             =   2735
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Tax:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Image Image8 
         Height          =   285
         Left            =   2760
         Top             =   2025
         Width           =   2955
         BackColor       =   16777215
         Size            =   "5212;503"
      End
      Begin MSForms.Image Image10 
         Height          =   285
         Left            =   2760
         Top             =   2370
         Width           =   2955
         BackColor       =   16777215
         Size            =   "5212;503"
      End
      Begin MSForms.TextBox txtDep 
         Height          =   285
         Left            =   2760
         TabIndex        =   0
         Top             =   990
         Width           =   1305
         VariousPropertyBits=   746604563
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "2302;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDepName 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   1335
         Width           =   6105
         VariousPropertyBits=   746604563
         MaxLength       =   50
         BorderStyle     =   1
         Size            =   "10769;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtShortName 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   1680
         Width           =   2955
         VariousPropertyBits=   746604563
         MaxLength       =   24
         BorderStyle     =   1
         Size            =   "5212;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmDepartments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CreateDepartment(DeptNo)
    If grdMajor.Rows = 1 Then
        lblType.Caption = "Major Department"
        lblParent.Caption = "None"
    Else
         lblParent.Caption = grdMajor.TextMatrix(grdMajor.Row, 1)
    End If
    
    txtDep.Text = ""
    txtDepName.Text = ""
    txtShortName.Text = ""
    For i = 1 To grdMinor.Rows - 1
        grdMinor.TextMatrix(i, 2) = "1"
    Next i
    chkZero.Value = 0
    txtDep.SetFocus
    frmMain.Toolbar1.Buttons(2).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    cmdChange.Enabled = False
End Sub

Private Sub cmdChange_Click()
    frmDepChange.Show vbModal
    Load_Departments
End Sub
Private Sub cmdUp2_Click()
    Select Case cmdUp2.Caption
        Case "5"
            picSeperate2.top = 0
            picTopBar2.top = picSeperate2.top + picSeperate2.Height - 30
            picFrame.top = picTopBar2.top + 60
            cmdUp2.top = picFrame.top
            picTree.top = picTopBar2.top + picTopBar2.Height - 10
            picTree.Height = picProducts(1).Height - picSeperate2.Height - picTopBar2.Height - 1390
            grdMajor.Height = picTree.Height - 450
            grdMinor.Height = picTree.Height - 450
            grdSub.Height = picTree.Height - 450
            txtDep.SetFocus
            cmdUp2.Caption = 6
        Case "6"
            picTopBar2.top = 3300
            picSeperate2.top = 3210
            cmdUp2.top = 3390
            picFrame.top = 3390
            picTree.top = 3900
            picTree.Height = 5325
            grdMajor.Height = 4515
            grdSub.Height = 4515
            grdMinor.Height = 4515
            cmdUp2.Caption = "5"
            txtDep.SetFocus
    End Select
End Sub

Private Sub Form_Activate()
    txtDep.SetFocus
    frmMain.Toolbar1.Buttons(2).Caption = "New"
End Sub

Private Sub txtShortName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If grdMajor.Rows = 1 And frmMain.Toolbar1.Buttons(2).Enabled = True Then
        KeyCode = 0
    End If
End Sub
Private Sub txtShortName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 45
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtShortName_LostFocus()
    On Error Resume Next
    txtShortName.Text = UCase(Left(txtShortName.Text, 1)) & Mid(txtShortName.Text, 2)
    On Error GoTo 0
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
Private Sub Load_Departments()
    On Error Resume Next
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(4).Enabled = False
    lblType.Caption = "Major Department"
    lblParent.Caption = "None"
    grdMinor.ColAlignment(2) = flexAlignCenterCenter
    grdMinor.ColDataType(2) = flexDTBoolean
    grdMajor.Rows = 1
    grdSub.Rows = 1
    grdMinor.Rows = 1
    ActiveReadServer "Select * from Locations order by Location_No"
    While Not rs.EOF
        grdMinor.Rows = grdMinor.Rows + 1
        grdMinor.Row = grdMinor.Rows - 1
        grdMinor.TextMatrix(grdMinor.Row, 0) = rs.Fields("Location_No")
        grdMinor.TextMatrix(grdMinor.Row, 1) = rs.Fields("Loc_name")
        rs.MoveNext
    Wend
    rs.Close
    If grdMinor.Rows > 0 Then grdMinor.Row = 1
    ActiveReadServer "Select * From Departments where Dept_Type=0 order by Department_No"
    i = 0
    While Not rs.EOF
        grdMajor.Rows = grdMajor.Rows + 1
        i = i + 1
        grdMajor.TextMatrix(i, 0) = rs.Fields("Department_No")
        grdMajor.TextMatrix(i, 1) = rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    If grdMajor.Rows > 1 Then
        grdMajor.Row = 1
        ActiveReadServer "Select * From Departments where Dept_Type=1 and Department_no= '" & grdMajor.TextMatrix(1, 0) & "' order by Department_No"
        i = 0
        While Not rs.EOF
            grdSub.Rows = grdSub.Rows + 1
            i = i + 1
            grdSub.TextMatrix(i, 0) = rs.Fields("Department_No")
            grdSub.TextMatrix(i, 1) = rs.Fields("Dept_Name")
            chkZero.Value = Val(rs.Fields("Zero_Price") & "")
            rs.MoveNext
        Wend
        rs.Close
        For i = 1 To grdMinor.Rows - 1
            grdMinor.TextMatrix(i, 2) = "0"
        Next i
    End If
    
    If grdMinor.Rows > 0 Then
        grdMinor.Row = 1
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    RecordCount = 0
    grdMajor.TextMatrix(0, 0) = "No."
    grdMajor.TextMatrix(0, 1) = "Description"
    grdSub.TextMatrix(0, 0) = "No."
    grdSub.TextMatrix(0, 1) = "Description"
    grdMinor.TextMatrix(0, 0) = "No."
    grdMinor.TextMatrix(0, 1) = "Description"
    grdMinor.TextMatrix(0, 2) = "Linked"
    grdMinor.ColWidth(0) = 600
    chkZero.Value = 0
    cmbTax1.Clear
    ActiveReadServer "Select * from Tax_Rates order by Tax_Type"
    While Not rs.EOF
        cmbTax1.AddItem rs.Fields("Tax_type") & " - " & rs.Fields("Tax_Rate") & "% " & rs.Fields("Description")
        rs.MoveNext
    Wend
    rs.Close
    If cmbTax1.ListCount > 1 Then
        cmbTax1.Text = cmbTax1.List(0)
    End If
    
    Load_Departments
    frmMain.stbBar.Panels(3) = ""
End Sub
Private Sub grdMajor_AfterSort(ByVal Col As Long, Order As Integer)
    grdMajor_RowColChange
End Sub
Private Sub grdMajor_Click()
    frmCodeRange.Visible = True
    fmOptions.Visible = False
    cmbTax1.Enabled = True
    lblTax1.Enabled = True
    chk_foreign_currency.Visible = False
    picMin.BackColor = &H8000000F
    picLoc.BackColor = &H8000000F
    picMaj.BackColor = &HFFC0C0
    lblType.Caption = "Major Department"
    If grdMajor.Rows > 1 Then
            frmMain.Toolbar1.Buttons(2).Enabled = True
            frmMain.Toolbar1.Buttons(4).Enabled = True
            frmMain.Toolbar1.Buttons(5).Enabled = True
    Else
            frmMain.Toolbar1.Buttons(2).Enabled = True
            frmMain.Toolbar1.Buttons(4).Enabled = False
            frmMain.Toolbar1.Buttons(5).Enabled = False
    End If
    DoEvents
    If grdMajor.Rows > 1 And picMaj.BackColor = &HFFC0C0 Then
        lblType.Caption = "Major Department"
        lblParent.Caption = "None"
        ActiveReadServer "Select *,(Select Description from Tax_Rates where Departments.Tax_Type=Tax_Rates.Tax_Type) as Tax_Desc From Departments where Dept_Type=0 and Department_No='" & grdMajor.TextMatrix(grdMajor.Row, 0) & "'"
        If rs.RecordCount > 0 Then
            txtDep.Text = rs.Fields("Department_No")
            txtDepName.Text = rs.Fields("Dept_Name")
            txtShortName.Text = rs.Fields("Short_Name")
            txtPluStart.Text = rs.Fields("Range_Start") & ""
            txtPluStop.Text = rs.Fields("Range_Stop") & ""
            If Not IsNull(rs.Fields("Tax_Type")) Then
                cmbTax1.Text = rs.Fields("Tax_Type") & " - " & rs.Fields("Sales_Tax") & "% " & rs.Fields("Tax_Desc")
            End If
        End If
        rs.Close
        
        For i = 1 To grdMinor.Rows - 1
            grdMinor.TextMatrix(i, 2) = "0"
        Next i
        ActiveReadServer "Select * from Department_links where Dept_No = '" & txtDep.Text & "'"
        While Not rs.EOF
            For i = 1 To grdMinor.Rows - 1
                If rs.Fields("Location_No") = grdMinor.TextMatrix(i, 0) Then
                    grdMinor.TextMatrix(i, 2) = "1"
                End If
            Next i
            rs.MoveNext
        Wend
        rs.Close
    End If
    If grdMajor.Rows > 1 Then cmdChange.Enabled = True
End Sub
Private Sub grdMajor_GotFocus()
    grdMajor_Click
End Sub
Private Sub grdMajor_RowColChange()
    frmCodeRange.Visible = True
    lblType.Caption = "Major Department"
    lblParent.Caption = "None"
    ActiveReadServer "Select *,(Select Description from Tax_Rates where Departments.Tax_Type=Tax_Rates.Tax_Type) as Tax_Desc From Departments where Dept_Type=0 and Department_No='" & grdMajor.TextMatrix(grdMajor.Row, 0) & "'"
    If rs.RecordCount > 0 Then
        txtDep.Text = rs.Fields("Department_No")
        txtDepName.Text = rs.Fields("Dept_Name")
        txtShortName.Text = rs.Fields("Short_Name")
        txtPluStart.Text = rs.Fields("Range_Start") & ""
        txtPluStop.Text = rs.Fields("Range_Stop") & ""
        If Not IsNull(rs.Fields("Tax_Type")) Then
            cmbTax1.Text = rs.Fields("Tax_Type") & " - " & rs.Fields("Sales_Tax") & "% " & rs.Fields("Tax_Desc")
        End If
    End If
    rs.Close
    grdSub.Rows = 1
    ActiveReadServer "Select * From Departments where Dept_Type=1 and Dept_Parent='" & txtDep.Text & "'"
    i = 0
    While Not rs.EOF
        i = i + 1
        grdSub.Rows = grdSub.Rows + 1
        grdSub.TextMatrix(i, 0) = rs.Fields("Department_No")
        grdSub.TextMatrix(i, 1) = rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    For i = 1 To grdMinor.Rows - 1
        grdMinor.TextMatrix(i, 2) = "0"
    Next i
    ActiveReadServer "Select * from Department_links where Dept_No = '" & txtDep.Text & "'"
    While Not rs.EOF
        For i = 1 To grdMinor.Rows - 1
            If rs.Fields("Location_No") = grdMinor.TextMatrix(i, 0) Then
                grdMinor.TextMatrix(i, 2) = "1"
            End If
        Next i
        rs.MoveNext
    Wend
    rs.Close
    If grdSub.Rows > 1 Then
        grdSub.Row = 1
        frmMain.Toolbar1.Buttons(5).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(5).Enabled = False
    End If
End Sub

Private Sub grdMinor_Click()
    If lblType.Caption = "Major Department" Then
        If grdMajor.Rows > 1 Then
            frmMain.Toolbar1.Buttons(2).Enabled = False
            frmMain.Toolbar1.Buttons(4).Enabled = True
        Else
            frmMain.Toolbar1.Buttons(2).Enabled = False
            frmMain.Toolbar1.Buttons(4).Enabled = False
        End If
    End If
    If lblType.Caption = "Sub Department" Then
        If grdSub.Rows > 1 Then
            frmMain.Toolbar1.Buttons(2).Enabled = False
            frmMain.Toolbar1.Buttons(4).Enabled = True
        Else
            frmMain.Toolbar1.Buttons(2).Enabled = False
            frmMain.Toolbar1.Buttons(4).Enabled = False
        End If
    End If
    frmMain.Toolbar1.Buttons(5).Enabled = False
End Sub
Private Sub grdMinor_GotFocus()
    grdMinor_Click
End Sub
Private Sub grdMinor_RowColChange()
    If grdMinor.Col = 2 Then
        grdMinor.Editable = flexEDKbdMouse
    Else
        grdMinor.Editable = flexEDNone
    End If
    frmMain.Toolbar1.Buttons(5).Enabled = False
End Sub


Private Sub grdSub_AfterSort(ByVal Col As Long, Order As Integer)
    grdSub_RowColChange
End Sub
Private Sub grdSub_Click()
    frmCodeRange.Visible = False
    fmOptions.Visible = True
    cmbTax1.Enabled = False
    lblTax1.Enabled = False
    chk_foreign_currency.Visible = True
    picMin.BackColor = &HFFC0C0
    picLoc.BackColor = &H8000000F
    picMaj.BackColor = &H8000000F
    lblType.Caption = "Sub Department"
    If grdSub.Rows = 1 Then
        txtDep.Text = ""
        txtDepName.Text = ""
        txtShortName.Text = ""
    End If
    If grdSub.Rows > 1 Then
            frmMain.Toolbar1.Buttons(2).Enabled = True
            frmMain.Toolbar1.Buttons(4).Enabled = True
            frmMain.Toolbar1.Buttons(5).Enabled = True
    Else
            frmMain.Toolbar1.Buttons(2).Enabled = True
            frmMain.Toolbar1.Buttons(4).Enabled = False
            frmMain.Toolbar1.Buttons(5).Enabled = False
    End If
    If grdSub.Rows > 1 And picMin.BackColor = &HFFC0C0 Then
        lblType.Caption = "Sub Department"
        lblParent.Caption = grdMajor.TextMatrix(grdMajor.Row, 1)
        ActiveReadServer "Select * From Departments where Dept_Type=1 and Department_No='" & grdSub.TextMatrix(grdSub.Row, 0) & "'"
        If rs.RecordCount > 0 Then
            txtDep.Text = rs.Fields("Department_No")
            txtDepName.Text = rs.Fields("Dept_Name")
            txtShortName.Text = rs.Fields("Short_Name")
            chkZero.Value = Val(rs.Fields("Zero_Price") & "")
            If rs.Fields("Report_forreign_curr") & "" = True Then chk_foreign_currency.Value = 1 Else chk_foreign_currency.Value = 0
        End If
        rs.Close
        
        For i = 1 To grdMinor.Rows - 1
            grdMinor.TextMatrix(i, 2) = "0"
        Next i
        ActiveReadServer "Select * from Department_links where Dept_No = '" & txtDep.Text & "'"
        While Not rs.EOF
            For i = 1 To grdMinor.Rows - 1
                If rs.Fields("Location_No") = grdMinor.TextMatrix(i, 0) Then
                    grdMinor.TextMatrix(i, 2) = "1"
                End If
            Next i
            rs.MoveNext
        Wend
        rs.Close
    End If
    If grdSub.Rows > 1 Then cmdChange.Enabled = True
End Sub
Private Sub grdSub_GotFocus()
    grdSub_Click
End Sub
Private Sub grdSub_RowColChange()
    frmCodeRange.Visible = False
    DoEvents
    If grdSub.Rows > 1 And picMin.BackColor = &HFFC0C0 Then
        lblType.Caption = "Sub Department"
        lblParent.Caption = grdMajor.TextMatrix(grdMajor.Row, 1)
        ActiveReadServer "Select * From Departments where Dept_Type=1 and Department_No='" & grdSub.TextMatrix(grdSub.Row, 0) & "'"
        If rs.RecordCount > 0 Then
            txtDep.Text = rs.Fields("Department_No")
            txtDepName.Text = rs.Fields("Dept_Name")
            txtShortName.Text = rs.Fields("Short_Name")
            chkZero.Value = Val(rs.Fields("Zero_Price") & "")
        End If
        rs.Close
        For i = 1 To grdMinor.Rows - 1
            grdMinor.TextMatrix(i, 2) = "0"
        Next i
        ActiveReadServer "Select * from Department_links where Dept_No = '" & txtDep.Text & "'"
        While Not rs.EOF
            For i = 1 To grdMinor.Rows - 1
                If rs.Fields("Location_No") = grdMinor.TextMatrix(i, 0) Then
                    grdMinor.TextMatrix(i, 2) = "1"
                End If
            Next i
            rs.MoveNext
        Wend
        rs.Close
        frmMain.Toolbar1.Buttons(5).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(5).Enabled = False
    End If
End Sub

Private Sub lblLinks_Click()
    grdMinor.SetFocus
    If picMaj.BackColor = &HFFC0C0 Then
        If grdMajor.Rows > 1 Then
            frmMain.Toolbar1.Buttons(2).Enabled = False
            frmMain.Toolbar1.Buttons(4).Enabled = True
        End If
    End If
    If picMin.BackColor = &HFFC0C0 Then
        If grdSub.Rows > 1 Then
            frmMain.Toolbar1.Buttons(2).Enabled = False
            frmMain.Toolbar1.Buttons(4).Enabled = True
        End If
    End If
    If grdMajor.Rows = 1 And grdMajor.Rows = 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub
Private Sub lblMaj_Click()
    picMin.BackColor = &H8000000F
    picLoc.BackColor = &H8000000F
    picMaj.BackColor = &HFFC0C0
    lblType.Caption = "Major Department"
    grdMajor.SetFocus
End Sub
Private Sub lblSub_Click()
    picMin.BackColor = &HFFC0C0
    picLoc.BackColor = &H8000000F
    picMaj.BackColor = &H8000000F
    lblType.Caption = "Sub Department"
    grdSub.SetFocus
End Sub


Private Sub txtDep_Change()
    If txtDep.Text = "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = False
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = True
    End If
    If txtDepName.Text = "" Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If picMaj.BackColor = &HFFC0C0 Then
        Select Case lblType.Caption
            Case "Major Department"
                If grdMajor.FindRow(txtDep.Text, 0, 0, 1, 1) <> -1 Then
                    grdMajor.Row = grdMajor.FindRow(txtDep.Text, 0, 0, 1, 1)
                End If
        End Select
    End If
    If picMin.BackColor = &HFFC0C0 And InStr(txtDep.Text, "-") = 0 Then
        Select Case lblType.Caption
            Case "Sub Department"
                If grdSub.FindRow(grdMajor.TextMatrix(grdMajor.Row, 0) & "-" & txtDep.Text, 0, 0, 1, 1) <> -1 Then
                    grdSub.Row = grdSub.FindRow(grdMajor.TextMatrix(grdMajor.Row, 0) & "-" & txtDep.Text, 0, 0, 1, 1)
                End If
        End Select
    Else
        txtDepName.Text = ""
        txtShortName.Text = ""
    End If
End Sub
Private Sub txtDep_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If grdMajor.Rows = 1 And frmMain.Toolbar1.Buttons(2).Enabled = True Then
        KeyCode = 0
    End If
End Sub

Private Sub txtDep_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If lblType.Caption = "Sub Department" Then
        Select Case KeyAscii
            Case 8, 48 To 57
            Case Else
                KeyAscii = 0
        End Select
    Else
        Select Case KeyAscii
            Case 8, 48 To 57, 65 To 90
            Case 97 To 122
                KeyAscii = KeyAscii - 32
            Case Else
                KeyAscii = 0
        End Select
    End If
End Sub
Private Sub txtDepName_Change()
    If txtDepName.Text = "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = False
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = True
    End If
    If txtDep.Text = "" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub txtDepName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If grdMajor.Rows = 1 And frmMain.Toolbar1.Buttons(2).Enabled = True Then
        KeyCode = 0
    End If
End Sub

Private Sub txtDepName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    
    Select Case KeyAscii
        Case 33 To 47, 58 To 64, 91 To 96, 123 To 127, 162 To 184, 247, 248, 191
            KeyAscii = 0
        
    End Select

End Sub
Private Sub txtDepName_LostFocus()
    On Error Resume Next
    txtDepName.Text = UCase(Left(txtDepName.Text, 1)) & Mid(txtDepName.Text, 2)
    If Len(txtDepName.Text) < 25 And txtShortName.Text = "" Then
        txtShortName.Text = txtDepName.Text
    End If
    On Error GoTo 0
End Sub
Private Sub txtPluStart_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtPluStart_LostFocus()
    If txtPluStart.Text = "" Then txtPluStart.Text = "0"
    If Val(txtPluStart.Text) > Val(txtPluStop.Text) Then
        MsgBox "Your Start Range cannot be bigger than the End Range", vbCritical, "HeroPOS"
        txtPluStart.Text = "0"
    End If
End Sub
Private Sub txtPluStop_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPluStop_LostFocus()
    If txtPluStop.Text = "" Then txtPluStop.Text = "0"
End Sub



