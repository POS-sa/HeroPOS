VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmCheck 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   11400
   ClientLeft      =   165
   ClientTop       =   -1890
   ClientWidth     =   14655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11400
   ScaleWidth      =   14655
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picRates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8160
      ScaleHeight     =   285
      ScaleWidth      =   2715
      TabIndex        =   82
      Top             =   3450
      Visible         =   0   'False
      Width           =   2745
      Begin VB.Label lblRatestring 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   150
         TabIndex        =   83
         Top             =   30
         Width           =   2325
      End
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   1470
      ScaleHeight     =   2925
      ScaleWidth      =   5355
      TabIndex        =   79
      Top             =   3930
      Visible         =   0   'False
      Width           =   5355
      Begin btButtonEx.ButtonEx cmdOk 
         Height          =   315
         Left            =   4110
         TabIndex        =   80
         Top             =   2460
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Appearance      =   3
         Caption         =   "Ok"
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
      Begin MSComCtl2.MonthView mthView 
         Height          =   2310
         Left            =   90
         TabIndex        =   81
         Top             =   90
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16239822
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxSelCount     =   365
         MonthColumns    =   2
         MonthBackColor  =   16777215
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   65994754
         TitleBackColor  =   16761281
         TrailingForeColor=   -2147483639
         CurrentDate     =   38701
      End
      Begin MSForms.Image Image6 
         Height          =   2805
         Index           =   2
         Left            =   60
         Top             =   60
         Width           =   5235
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "9234;4948"
      End
      Begin MSForms.Image Image5 
         Height          =   2925
         Index           =   16
         Left            =   30
         Top             =   0
         Width           =   5325
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "9393;5159"
      End
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   14595
      TabIndex        =   2
      Top             =   6600
      Width           =   14595
      Begin btButtonEx.ButtonEx cmdUp 
         Height          =   465
         Left            =   13080
         TabIndex        =   3
         Top             =   210
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
      Begin btButtonEx.ButtonEx ButtonEx1 
         Height          =   375
         Left            =   6840
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   300
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   8421504
         Caption         =   "¦"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdAccount 
         Height          =   465
         Left            =   10740
         TabIndex        =   87
         Top             =   210
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "View Room Account"
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
      Begin MSForms.ComboBox cmb1 
         Height          =   375
         Left            =   7320
         TabIndex        =   78
         Top             =   300
         Width           =   3375
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "5953;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Arial Narrow"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblDate 
         Height          =   315
         Left            =   4500
         TabIndex        =   77
         Top             =   390
         Width           =   2235
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "1 Feb 2006 to 13 Feb 2006"
         Size            =   "3942;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image picSeperate1 
         Height          =   105
         Left            =   -30
         Top             =   0
         Width           =   13815
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "24368;185"
      End
      Begin VB.Label Label5 
         Caption         =   "Guest List..."
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
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   210
         Width           =   3135
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   2
         Left            =   3870
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   1
         Left            =   3540
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   0
         Left            =   3210
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image6 
         Height          =   90
         Index           =   0
         Left            =   150
         Top             =   570
         Width           =   3015
         BackColor       =   16761024
         Size            =   "5318;159"
      End
      Begin MSForms.Image Image1 
         Height          =   375
         Left            =   4440
         Top             =   300
         Width           =   2385
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "4207;661"
      End
      Begin MSForms.Image Image7 
         Height          =   675
         Left            =   -1260
         Top             =   90
         Width           =   15030
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "26511;1191"
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdGuest 
      Height          =   2010
      Left            =   -60
      TabIndex        =   0
      Top             =   7350
      Width           =   13785
      _cx             =   24315
      _cy             =   3545
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
      BackColorSel    =   14408671
      ForeColorSel    =   8388608
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
      Rows            =   100
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCheck.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      WordWrap        =   -1  'True
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
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   1
         Top             =   11700
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   315
      Left            =   9240
      TabIndex        =   23
      Top             =   6810
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Cancel"
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
   Begin btButtonEx.ButtonEx cmdAccept 
      Height          =   315
      Left            =   7530
      TabIndex        =   28
      Top             =   6810
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Accept"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":0078
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":0FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":1500
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":1A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":1F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":2466
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":2988
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":2EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":33CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":38EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":3E10
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":4332
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":4854
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":4D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":5298
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":57BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":5CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":61FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":6720
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":6C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":7164
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":7686
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":7BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":80CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":85EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":8B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheck.frx":9030
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdDeposit 
      Height          =   315
      Left            =   5820
      TabIndex        =   33
      Top             =   6810
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Receive Deposit"
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
   Begin BTNENHLib4.BtnEnh fmType 
      Height          =   1275
      Left            =   5070
      TabIndex        =   32
      Top             =   5280
      Width           =   2955
      _Version        =   524298
      _ExtentX        =   5212
      _ExtentY        =   2249
      _StockProps     =   66
      Caption         =   "Provisional Booking"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
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
      Surface         =   3
      PictureTranspColor=   192
      BackColorContainer=   14215660
      ShadowColor     =   16777215
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      ForeColorDisabled=   12640511
      UserData        =   0.1
      textCaption     =   "frmCheck.frx":9552
      textLT          =   "frmCheck.frx":95D8
      textCT          =   "frmCheck.frx":95F0
      textRT          =   "frmCheck.frx":9608
      textLM          =   "frmCheck.frx":9620
      textRM          =   "frmCheck.frx":9638
      textLB          =   "frmCheck.frx":9650
      textCB          =   "frmCheck.frx":9668
      textRB          =   "frmCheck.frx":9680
      colorBack       =   "frmCheck.frx":9698
      colorIntern     =   "frmCheck.frx":96C2
      colorMO         =   "frmCheck.frx":96EC
      colorFocus      =   "frmCheck.frx":9716
      colorDisabled   =   "frmCheck.frx":9740
      colorPressed    =   "frmCheck.frx":976A
      HollowFrame     =   -1  'True
      LightDirection  =   8
   End
   Begin RichTextLib.RichTextBox txtRemarks 
      Height          =   1485
      Left            =   8310
      TabIndex        =   31
      Top             =   4980
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   2619
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmCheck.frx":9794
   End
   Begin btButtonEx.ButtonEx cmdCondition 
      Height          =   345
      Left            =   2880
      TabIndex        =   30
      Top             =   120
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Conditions of Stay..."
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
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   1155
      Left            =   1740
      TabIndex        =   29
      Top             =   5310
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2037
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmCheck.frx":9816
   End
   Begin MSComCtl2.DTPicker DTArrival 
      Height          =   345
      Left            =   1620
      TabIndex        =   27
      Top             =   870
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65994755
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTDeparture 
      Height          =   345
      Left            =   1620
      TabIndex        =   26
      Top             =   1260
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65994755
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTDepTime 
      Height          =   345
      Left            =   3660
      TabIndex        =   25
      Top             =   1260
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   65994754
      CurrentDate     =   38862.4583333333
   End
   Begin MSComCtl2.DTPicker DTArrTime 
      Height          =   345
      Left            =   3660
      TabIndex        =   24
      Top             =   870
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   65994754
      CurrentDate     =   38862.5833333333
   End
   Begin VB.TextBox txtContactNo 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   8310
      TabIndex        =   6
      Top             =   3120
      Width           =   2565
   End
   Begin VB.TextBox txtContact 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   8310
      TabIndex        =   7
      Top             =   2760
      Width           =   2565
   End
   Begin VB.TextBox txt5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   10140
      TabIndex        =   8
      Text            =   "0"
      Top             =   1710
      Width           =   735
   End
   Begin VB.TextBox txtAdults 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   8310
      TabIndex        =   9
      Text            =   "0"
      Top             =   1710
      Width           =   735
   End
   Begin VB.TextBox txt0to5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   8310
      TabIndex        =   10
      Text            =   "0"
      Top             =   2070
      Width           =   735
   End
   Begin VB.TextBox txt12to16 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   10140
      TabIndex        =   11
      Text            =   "0"
      Top             =   2070
      Width           =   735
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   8310
      TabIndex        =   12
      Top             =   1290
      Width           =   735
   End
   Begin VB.TextBox txtNights 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   13
      Top             =   1680
      Width           =   1035
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   14
      Top             =   2400
      Width           =   3165
   End
   Begin VB.TextBox txtFax 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   15
      Top             =   4920
      Width           =   3075
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   16
      Top             =   4560
      Width           =   3165
   End
   Begin VB.TextBox txtTel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   17
      Top             =   4200
      Width           =   3165
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   18
      Top             =   3840
      Width           =   4275
   End
   Begin VB.TextBox txtVehReg 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   19
      Top             =   3480
      Width           =   2595
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   20
      Top             =   3120
      Width           =   2595
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   21
      Top             =   2760
      Width           =   4275
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1770
      TabIndex        =   22
      Top             =   540
      Width           =   4185
   End
   Begin btButtonEx.ButtonEx cmdRate 
      Height          =   300
      Left            =   6720
      TabIndex        =   84
      Top             =   3450
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Applied Rates..."
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
      ShowFocus       =   0
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rate:"
      Height          =   255
      Left            =   6690
      TabIndex        =   86
      Top             =   3870
      Width           =   1395
   End
   Begin VB.Label lblTotRate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0.00"
      Height          =   195
      Left            =   8340
      TabIndex        =   85
      Top             =   3885
      Width           =   1215
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   75
      Top             =   1320
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "*Departure Date:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   74
      Top             =   930
      Width           =   1395
      VariousPropertyBits=   8388627
      Caption         =   "*Arrival Date:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Index           =   12
      Left            =   180
      TabIndex        =   73
      Top             =   570
      Width           =   1395
      VariousPropertyBits=   8388627
      Caption         =   " Room Description:"
      Size            =   "2461;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   72
      Top             =   210
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Room Number:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Title:"
      Height          =   255
      Left            =   150
      TabIndex        =   71
      Top             =   2100
      Width           =   1395
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nights:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   70
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*First Name:"
      Height          =   255
      Left            =   150
      TabIndex        =   69
      Top             =   2430
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Last Name:"
      Height          =   255
      Left            =   150
      TabIndex        =   68
      Top             =   2790
      Width           =   1395
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID No:"
      Height          =   255
      Left            =   150
      TabIndex        =   67
      Top             =   3150
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Veh Reg No:"
      Height          =   255
      Left            =   150
      TabIndex        =   66
      Top             =   3510
      Width           =   1395
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   150
      TabIndex        =   65
      Top             =   3870
      Width           =   1395
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Tel No:"
      Height          =   255
      Left            =   150
      TabIndex        =   64
      Top             =   4230
      Width           =   1395
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No:"
      Height          =   255
      Left            =   150
      TabIndex        =   61
      Top             =   4590
      Width           =   1395
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Required fields = *"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   6900
      Width           =   1545
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Address:"
      Height          =   195
      Left            =   330
      TabIndex        =   38
      Top             =   5250
      Width           =   1215
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      Height          =   165
      Left            =   1290
      TabIndex        =   36
      Top             =   6840
      Width           =   1335
   End
   Begin MSForms.Image Image2 
      Height          =   315
      Left            =   2700
      Top             =   6810
      Width           =   2895
      BackColor       =   16777215
      Size            =   "5106;556"
   End
   Begin VB.Label lblUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2760
      TabIndex        =   35
      Top             =   6870
      Width           =   2745
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Fax No:"
      Height          =   255
      Left            =   150
      TabIndex        =   34
      Top             =   4950
      Width           =   1395
   End
   Begin MSForms.Label Label2 
      Height          =   555
      Index           =   10
      Left            =   3840
      TabIndex        =   5
      Top             =   6840
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Departure Date:"
      Size            =   "2461;979"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ComboBox cmbTitle 
      Height          =   315
      Left            =   1620
      TabIndex        =   63
      Tag             =   "Up"
      Top             =   2010
      Width           =   1245
      VariousPropertyBits=   746604569
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "2196;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbRoomNo 
      Height          =   345
      Left            =   1620
      TabIndex        =   62
      Tag             =   "Up"
      Top             =   120
      Width           =   1215
      VariousPropertyBits=   746604569
      MaxLength       =   4
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "2143;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Country of Origin:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   60
      Top             =   210
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   345
      Left            =   8160
      TabIndex        =   59
      Tag             =   "Up"
      Top             =   120
      Width           =   2745
      VariousPropertyBits=   746604569
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4842;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Province:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   58
      Top             =   570
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbProvince 
      Height          =   315
      Left            =   8160
      TabIndex        =   57
      Tag             =   "Up"
      Top             =   510
      Width           =   2745
      VariousPropertyBits=   746604569
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "Western Cape"
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City/Town:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   56
      Top             =   930
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbCity 
      Height          =   345
      Left            =   8160
      TabIndex        =   55
      Tag             =   "Up"
      Top             =   870
      Width           =   2745
      VariousPropertyBits=   746604569
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4842;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "Cape Town"
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   54
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Adults:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   53
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5-12 yrs:"
      Enabled         =   0   'False
      Height          =   225
      Left            =   8520
      TabIndex        =   52
      Top             =   1740
      Width           =   1395
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   51
      Top             =   2790
      Width           =   1395
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0-5 yrs:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7380
      TabIndex        =   50
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "12-16 yrs:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9210
      TabIndex        =   49
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Booked by:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   48
      Top             =   2430
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbBooked 
      Height          =   315
      Left            =   8160
      TabIndex        =   47
      Tag             =   "Up"
      Top             =   2370
      Width           =   2745
      VariousPropertyBits=   746604569
      MaxLength       =   25
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   46
      Top             =   3150
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbRate 
      Height          =   315
      Left            =   8160
      TabIndex        =   45
      Tag             =   "Up"
      Top             =   3450
      Width           =   2745
      VariousPropertyBits=   746604569
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business Source:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6690
      TabIndex        =   44
      Top             =   4230
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbBusiness 
      Height          =   315
      Left            =   8160
      TabIndex        =   43
      Tag             =   "Up"
      Top             =   4170
      Width           =   2745
      VariousPropertyBits=   746604569
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Responsibility:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6150
      TabIndex        =   42
      Top             =   4590
      Width           =   1935
   End
   Begin MSForms.ComboBox cmbPay 
      Height          =   315
      Left            =   8160
      TabIndex        =   41
      Tag             =   "Up"
      Top             =   4530
      Width           =   2745
      VariousPropertyBits=   746604569
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6150
      TabIndex        =   40
      Top             =   4950
      Width           =   1935
   End
   Begin MSForms.Image Image3 
      Height          =   1305
      Left            =   1620
      Top             =   5250
      Width           =   3345
      BorderColor     =   0
      BackColor       =   16777215
      Size            =   "5900;2302"
      VariousPropertyBits=   25
   End
   Begin VB.Label lblResNo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   5430
      TabIndex        =   37
      Top             =   900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Image Image4 
      Height          =   1665
      Index           =   3
      Left            =   8160
      Top             =   4890
      Width           =   2745
      BorderColor     =   0
      BackColor       =   16777215
      Size            =   "4842;2937"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   0
      Left            =   1620
      Top             =   510
      Width           =   4455
      BackColor       =   16777215
      Size            =   "7858;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image6 
      Height          =   315
      Index           =   1
      Left            =   1620
      Top             =   1650
      Width           =   1245
      BackColor       =   16777215
      Size            =   "2196;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   1
      Left            =   1620
      Top             =   2370
      Width           =   3345
      BackColor       =   16777215
      Size            =   "5900;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   2
      Left            =   1620
      Top             =   2730
      Width           =   4455
      BackColor       =   16777215
      Size            =   "7858;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   3
      Left            =   1620
      Top             =   3090
      Width           =   2775
      BackColor       =   16777215
      Size            =   "4895;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   4
      Left            =   1620
      Top             =   3450
      Width           =   2775
      BackColor       =   16777215
      Size            =   "4895;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   5
      Left            =   1620
      Top             =   3810
      Width           =   4455
      BackColor       =   16777215
      Size            =   "7858;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   6
      Left            =   1620
      Top             =   4170
      Width           =   3345
      BackColor       =   16777215
      Size            =   "5900;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   7
      Left            =   1620
      Top             =   4530
      Width           =   3345
      BackColor       =   16777215
      Size            =   "5900;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   8
      Left            =   1620
      Top             =   4890
      Width           =   3345
      BackColor       =   16777215
      Size            =   "5900;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   9
      Left            =   8160
      Top             =   1260
      Width           =   915
      BackColor       =   16777215
      Size            =   "1614;609"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   10
      Left            =   9990
      Top             =   2010
      Width           =   915
      BackColor       =   16777215
      Size            =   "1614;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   11
      Left            =   8160
      Top             =   1650
      Width           =   915
      BackColor       =   16777215
      Size            =   "1614;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   12
      Left            =   8160
      Top             =   2010
      Width           =   915
      BackColor       =   16777215
      Size            =   "1614;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   13
      Left            =   9990
      Top             =   1650
      Width           =   915
      BackColor       =   16777215
      Size            =   "1614;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   14
      Left            =   8160
      Top             =   2730
      Width           =   2745
      BackColor       =   16777215
      Size            =   "4842;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   15
      Left            =   8160
      Top             =   3090
      Width           =   2745
      BackColor       =   16777215
      Size            =   "4842;556"
      VariousPropertyBits=   25
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   17
      Left            =   8160
      Top             =   3810
      Width           =   1605
      BackColor       =   16777215
      Size            =   "2831;556"
      VariousPropertyBits=   25
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Load_Res()
    If grdGuest.ValueMatrix(grdGuest.Row, 8) = 0 Then Exit Sub
    ActiveReadServer "Select * from Reservations where Res_no = " & grdGuest.TextMatrix(grdGuest.Row, 8)
    lblResNo.Caption = Res_No
    If rs.RecordCount > 0 Then
        Select Case rs.Fields("Res_Type")
            Case 0
                fmType.ForeColor = &HC0C0&
                fmType.Caption = "Provisional Booking"
            Case 1
                fmType.ForeColor = &H8000&
                fmType.Caption = "Confirmed Booking"
            Case 2
                fmType.ForeColor = &HC00000
                fmType.Caption = "Guest Checked In"
            Case 3
                fmType.ForeColor = &HC0&
                fmType.Caption = "Guest Checked Out"
        End Select
    End If
    cmbRoomNo.Text = grdGuest.TextMatrix(grdGuest.Row, 0)
    txtDescription.Text = grdGuest.TextMatrix(grdGuest.Row, 1)
    DTArrival.Value = rs.Fields("Arrive_Date") & ""
    DTDeparture.Value = rs.Fields("Depart_Date") & ""
    cmbTitle.Text = rs.Fields("Title") & ""
    txtFName.Text = rs.Fields("First_Name") & ""
    txtLName.Text = rs.Fields("Last_Name") & ""
    txtId.Text = rs.Fields("ID_No") & ""
    txtVehReg.Text = rs.Fields("Title") & ""
    txtEmail.Text = rs.Fields("Vehicle_No") & ""
    txttel.Text = rs.Fields("Tel_No") & ""
    txtMobile.Text = rs.Fields("Cell_No") & ""
    txtFax.Text = rs.Fields("Fax_No") & ""
    txtAddress.Text = rs.Fields("Address") & ""
    txtCode.Text = rs.Fields("Post_Code") & ""
    txtAdults.Text = rs.Fields("Adults") & ""
    cmbCountry.Text = rs.Fields("Country") & ""
    cmbProvince = rs.Fields("Province") & ""
    cmbCity = rs.Fields("City") & ""
    txt5.Text = Val(rs.Fields("Kid5to12") & "")
    txt0to5.Text = Val(rs.Fields("Kid0to5") & "")
    txt12to16.Text = Val(rs.Fields("Kids12to16") & "")
    txtContact.Text = rs.Fields("Contact_Person") & ""
    txtContactNo.Text = rs.Fields("Contact_no") & ""
    txtRemarks.Text = rs.Fields("Remarks") & ""
    cmbBusiness.Text = rs.Fields("Source") & ""
    cmbPay.Text = rs.Fields("Payment") & ""
    cmbBooked.Text = rs.Fields("Booked_By") & ""
    For i = 0 To cmbRate.ListCount - 1
        If Val(Mid(cmbRate.List(i), 1, InStr(cmbRate.List(i), "-") - 1)) = rs.Fields("Rate_Type") Then
            cmbRate.ListIndex = rs.Fields("Rate_Type") - 1
            Exit For
        End If
    Next i
    rs.Close
    Calc_Rate
End Sub
Private Sub Load_Res_List()
    If Right(Str(Time_Stop), 2) = "AM" And mthViewStart.Value = mthViewEnd.Value Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    grdGuest.Rows = 1
    Select Case cmb1.Text
        Case "<All Reservations>"
            ActiveReadServer "Select * from Res_View " & _
            " where (Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Provisional Bookings"
            ActiveReadServer "Select * from Res_View" & _
            " where Res_Type = 0 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Confirmed Bookings"
            ActiveReadServer "Select * from Res_View" & _
            " where Res_Type = 1 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Guests Checked In"
            ActiveReadServer "Select * from Res_View" & _
            " where Res_Type = 2 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Guests Checked Out"
            ActiveReadServer "Select * from Res_View" & _
            " where Res_Type = 3 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Guests Arriving"
            ActiveReadServer "Select * from Res_View" & _
            " where (Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Guests Departing"
            ActiveReadServer "Select * from Res_View" & _
            " where (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Depart_Date"
    End Select
    
    While Not rs.EOF
        grdGuest.Rows = grdGuest.Rows + 1
        grdGuest.TextMatrix(grdGuest.Rows - 1, 0) = rs.Fields("Room_No")
        grdGuest.TextMatrix(grdGuest.Rows - 1, 1) = rs.Fields("Description")
        grdGuest.TextMatrix(grdGuest.Rows - 1, 2) = Format(rs.Fields("Arrive_Date"), "dd MMM yyyy ddd")
        grdGuest.TextMatrix(grdGuest.Rows - 1, 3) = Format(rs.Fields("Depart_Date"), "dd MMM yyyy ddd")
        grdGuest.TextMatrix(grdGuest.Rows - 1, 4) = rs.Fields("Guest_Name")
        grdGuest.TextMatrix(grdGuest.Rows - 1, 5) = rs.Fields("Tel_No")
        grdGuest.TextMatrix(grdGuest.Rows - 1, 6) = rs.Fields("Nights")
        grdGuest.TextMatrix(grdGuest.Rows - 1, 7) = Format(Val(rs.Fields("Balance") & ""), "0.00")
        grdGuest.TextMatrix(grdGuest.Rows - 1, 8) = rs.Fields("Res_No")
        Select Case rs.Fields("Res_Type")
            Case 0: grdGuest.Cell(flexcpBackColor, grdGuest.Rows - 1, 0, grdGuest.Rows - 1, 8) = &HC0FFFF
            Case 1: grdGuest.Cell(flexcpBackColor, grdGuest.Rows - 1, 0, grdGuest.Rows - 1, 8) = &HC0FFC0
            Case 2: grdGuest.Cell(flexcpBackColor, grdGuest.Rows - 1, 0, grdGuest.Rows - 1, 8) = &HFDE0DF
            Case 3: grdGuest.Cell(flexcpBackColor, grdGuest.Rows - 1, 0, grdGuest.Rows - 1, 8) = &HC0C0FF
        End Select
        rs.MoveNext
    Wend
    rs.Close
    If grdGuest.Rows > 1 Then grdGuest.Row = 1
End Sub
Private Sub ButtonEx1_Click()
    Select Case ButtonEx1.Value
        Case 0
            picDate.Visible = True
        Case 1
            picDate.Visible = False
            If picDate.Visible = False Then Load_Res_List
    End Select
End Sub
Private Sub cmb1_Change()
    If grdGuest.Tag = "" Then
        Load_Res_List
    End If
End Sub
Private Sub cmb1_GotFocus()
    picDate.Visible = False
End Sub
Private Sub cmdAccount_Click()
    Load frmAccount
    TillData.Res_No = grdGuest.TextMatrix(grdGuest.Row, 8)
    frmAccount.Tag = "Room"
    frmAccount.Show vbModal
End Sub
Private Sub cmdOk_Click()
    picDate.Visible = False
    If picDate.Visible = False Then Load_Res_List
End Sub

Private Sub cmdRate_Click()
    Load frmRates
    frmRates.Tag = "Check1"
    frmRates.Show vbModal
End Sub
Private Sub cmdUp_Click()
    Select Case cmdUp.Caption
        Case "5"
            grdGuest.SetFocus
            picHead.top = 0
            grdGuest.top = picHead.top + picHead.Height
            grdGuest.Height = frmCheck.Height - picHead.Height - 120
            DoEvents
            cmdUp.Caption = 6
        Case "6"
            cmdUp.Caption = "5"
            picHead.top = 6600
            grdGuest.top = 7350
            grdGuest.Height = 2010
            DoEvents
    End Select
End Sub

Private Sub Form_Activate()
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Check"
End Sub
Private Sub Calc_Rate()
    picRates.Visible = False
    cmbRate.Enabled = True
    cmdRate.Enabled = False
    lblRatestring = ""
    If Val(txtAdults.Text) <> 0 Then
        lblTotRate.Caption = Format(Val(Mid(cmbRate.Text, InStr(cmbRate.Text, ">") + 1)), "0.00")
        lblRatestring.Caption = Trim(Mid(cmbRate.Text, 1, InStr(cmbRate.Text, "-") - 1))
    Else
        lblTotRate.Caption = "0.00"
        lblRatestring.Caption = ""
    End If
    If Val(txt5.Text) <> 0 Then
        ActiveReadServer2 "Select * from Room_Rates where Condition = 'Children 5 to 12'"
        If rs2.RecordCount > 0 Then
            lblRatestring.Caption = lblRatestring.Caption & "-" & rs2.Fields("Rate_Type")
            lblTotRate.Caption = Format(Val(lblTotRate.Caption) + (Val(txt5.Text) * rs2.Fields("Room_Rate")), "0.00")
        End If
        rs2.Close
    End If
    If Val(txt0to5.Text) <> 0 Then
        ActiveReadServer2 "Select * from Room_Rates where Condition = 'Children under 5'"
        If rs2.RecordCount > 0 Then
            lblRatestring.Caption = lblRatestring.Caption & "-" & rs2.Fields("Rate_Type")
            lblTotRate.Caption = Format(Val(lblTotRate.Caption) + (Val(txt0to5.Text) * rs2.Fields("Room_Rate")), "0.00")
        End If
        rs2.Close
    End If
    If Val(txt12to16.Text) <> 0 Then
        ActiveReadServer2 "Select * from Room_Rates where Condition = 'Children 12 to 15'"
        If rs2.RecordCount > 0 Then
            lblRatestring.Caption = lblRatestring.Caption & "-" & rs2.Fields("Rate_Type")
            lblTotRate.Caption = Format(Val(lblTotRate.Caption) + (Val(txt12to16.Text) * rs2.Fields("Room_Rate")), "0.00")
        End If
        rs2.Close
    End If
    If InStr(lblRatestring.Caption, "-") <> 0 Then
        cmdRate.Enabled = True
        cmbRate.Enabled = False
        picRates.Visible = True
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            KeyCode = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
            frmSplash.Show
            frmMain.Hide
    End Select
End Sub
Private Sub Form_Load()
    On Error Resume Next
    grdGuest.Tag = "1"
    frmMain.Toolbar1.Buttons(4).Enabled = False
    mthViewStart.Value = Date
    mthViewEnd.Value = Date
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
    cmb1.Clear
    cmb1.AddItem "<All Reservations>"
    cmb1.AddItem "Provisional Bookings"
    cmb1.AddItem "Confirmed Bookings"
    cmb1.AddItem "Guests Checked In"
    cmb1.AddItem "Guests Checked Out"
    cmb1.AddItem "Guests Arriving"
    cmb1.AddItem "Guests Departing"
    cmb1.Text = "<All Reservations>"
    grdGuest.Rows = 1
    grdGuest.Cols = 9
    grdGuest.ColHidden(8) = True
    grdGuest.TextMatrix(0, 0) = "Room Number"
    grdGuest.TextMatrix(0, 1) = "Room Description"
    grdGuest.TextMatrix(0, 2) = "Arrival Date"
    grdGuest.TextMatrix(0, 3) = "Departure Date"
    grdGuest.TextMatrix(0, 4) = "Guest Name"
    grdGuest.TextMatrix(0, 5) = "Tel. Number"
    grdGuest.TextMatrix(0, 6) = "Nights"
    grdGuest.TextMatrix(0, 7) = "Balance "
    grdGuest.ColWidth(0) = grdGuest.Width * 0.05
    grdGuest.ColWidth(1) = grdGuest.Width * 0.18
    grdGuest.ColWidth(2) = grdGuest.Width * 0.15
    grdGuest.ColWidth(3) = grdGuest.Width * 0.15
    grdGuest.ColWidth(4) = grdGuest.Width * 0.2
    grdGuest.ColWidth(5) = grdGuest.Width * 0.1
    grdGuest.ColWidth(6) = grdGuest.Width * 0.07
    grdGuest.ColAlignment(0) = flexAlignLeftCenter
    grdGuest.ColAlignment(1) = flexAlignLeftCenter
    grdGuest.ColAlignment(2) = flexAlignLeftCenter
    grdGuest.ColAlignment(3) = flexAlignLeftCenter
    grdGuest.ColAlignment(4) = flexAlignLeftCenter
    grdGuest.ColAlignment(5) = flexAlignLeftCenter
    grdGuest.ColAlignment(6) = flexAlignRightCenter
    grdGuest.ColAlignment(7) = flexAlignRightCenter
  
    cmbBusiness.Clear
        cmbBusiness.AddItem "Guests"
        cmbBusiness.AddItem "Travel Agent"
        cmbBusiness.AddItem "Advertisment"
        cmbBusiness.Text = "Guests"
        cmbBooked.Clear
        cmbBooked.AddItem "Guests"
        cmbBooked.AddItem "Travel Agent"
        cmbBooked.Text = "Guests"
        cmbPay.Clear
        cmbPay.AddItem "Guests"
        cmbPay.AddItem "Travel Agent"
        cmbPay.Text = "Guests"
        cmbRate.Clear
        ActiveReadServer1 "Select * from Room_Rates  where Active=1 order by Rate_Type"
        While Not rs1.EOF
            cmbRate.AddItem rs1.Fields("Rate_Type") & " - " & rs1.Fields("Description") & " > " & Format(rs1.Fields("Room_Rate"), "0.00")
            rs1.MoveNext
        Wend
        rs1.MoveFirst
        cmbRate.Text = rs1.Fields("Rate_Type") & " - " & rs1.Fields("Description") & " > " & Format(rs1.Fields("Room_Rate"), "0.00")
        rs1.Close
        ActiveReadServer "Select Room_Rate from Rooms where Room_No = " & frmRes.grdRes.ValueMatrix(frmRes.grdRes.Row, 1)
        If rs.RecordCount > 0 Then
            For i = 0 To cmbRate.ListCount - 1
                If Val(Mid(cmbRate.List(i, 0), 1, InStr(cmbRate.List(i, 0), "-") - 1)) = Val(rs.Fields("Room_Rate") & "") Then
                    cmbRate.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        rs.Close
        cmbCountry.Clear
        cmbCountry.AddItem "Angola"
        cmbCountry.AddItem "Botswana"
        cmbCountry.AddItem "Malawi"
        cmbCountry.AddItem "Mosambique"
        cmbCountry.AddItem "Namibia"
        cmbCountry.AddItem "South Africa"
        cmbCountry.AddItem "Zimbabwe"
        cmbCountry.Text = "South Africa"
        cmbCity.Clear
        cmbCity.AddItem "Bloemfontein"
        cmbCity.AddItem "Cape Town"
        cmbCity.AddItem "Durban"
        cmbCity.AddItem "East London"
        cmbCity.AddItem "Johannesburg"
        cmbCity.AddItem "Port Elizabeth"
        cmbCity.AddItem "Pretoria"
        cmbCity.AddItem "Polokwane"
        cmbCity.AddItem "Windhoek"
        cmbCity.AddItem "Gaborone"
        cmbCity.AddItem "Maputo"
        cmbTitle.Clear
        cmbTitle.AddItem "Dr."
        cmbTitle.AddItem "Prof."
        cmbTitle.AddItem "Miss."
        cmbTitle.AddItem "Mr."
        cmbTitle.AddItem "Mrs."
        cmbTitle.AddItem "Ms."
        cmbTitle.Text = "Mr."
        cmbProvince.Clear
        ActiveReadServer "Select * from Regions"
        While Not rs.EOF
            cmbProvince.AddItem rs.Fields("Region_Name")
            rs.MoveNext
        Wend
        rs.Close
        cmbRoomNo.Clear
        ActiveReadServer "Select Room_No,Description from Rooms order by Room_No"
        While Not rs.EOF
            cmbRoomNo.AddItem rs.Fields("Room_No")
            rs.MoveNext
        Wend
        rs.Close
        grdGuest.Tag = ""
        Load_Res_List
        On Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
End Sub

Private Sub grdGuest_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If grdGuest.Tag = "" Then
        If OldRow <> NewRow Then Load_Res
    End If
End Sub

Private Sub grdGuest_AfterSort(ByVal Col As Long, Order As Integer)
    Load_Res
End Sub

Private Sub grdGuest_Click()
    picDate.Visible = False
End Sub

Private Sub mthView_LostFocus()
    DoEvents
    If picDate.Visible = False Then Load_Res_List
End Sub
Private Sub mthView_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
End Sub

