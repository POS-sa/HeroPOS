VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSales1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0FFC0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSales1.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   200
      Left            =   3660
      Top             =   270
   End
   Begin BTNENHLib4.BtnEnh cmdLogoff 
      Height          =   1755
      Left            =   360
      TabIndex        =   39
      Top             =   600
      Width           =   1245
      _Version        =   524298
      _ExtentX        =   2196
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Log Off"
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
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":10089
      textLT          =   "frmSales1.frx":100F7
      textCT          =   "frmSales1.frx":1010F
      textRT          =   "frmSales1.frx":10127
      textLM          =   "frmSales1.frx":1013F
      textRM          =   "frmSales1.frx":10157
      textLB          =   "frmSales1.frx":1016F
      textCB          =   "frmSales1.frx":10187
      textRB          =   "frmSales1.frx":1019F
      colorBack       =   "frmSales1.frx":101B7
      colorIntern     =   "frmSales1.frx":101E1
      colorMO         =   "frmSales1.frx":1020B
      colorFocus      =   "frmSales1.frx":10235
      colorDisabled   =   "frmSales1.frx":1025F
      colorPressed    =   "frmSales1.frx":10289
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1215
      Index           =   1
      Left            =   1590
      TabIndex        =   40
      Top             =   2340
      Width           =   1965
      _Version        =   524298
      _ExtentX        =   3466
      _ExtentY        =   2143
      _StockProps     =   66
      Caption         =   "Transfer Items"
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
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":102B3
      textLT          =   "frmSales1.frx":1032F
      textCT          =   "frmSales1.frx":10347
      textRT          =   "frmSales1.frx":1035F
      textLM          =   "frmSales1.frx":10377
      textRM          =   "frmSales1.frx":1038F
      textLB          =   "frmSales1.frx":103A7
      textCB          =   "frmSales1.frx":103BF
      textRB          =   "frmSales1.frx":103D7
      colorBack       =   "frmSales1.frx":103EF
      colorIntern     =   "frmSales1.frx":10419
      colorMO         =   "frmSales1.frx":10443
      colorFocus      =   "frmSales1.frx":1046D
      colorDisabled   =   "frmSales1.frx":10497
      colorPressed    =   "frmSales1.frx":104C1
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1235
      Index           =   18
      Left            =   360
      TabIndex        =   38
      Top             =   2340
      Width           =   1245
      _Version        =   524298
      _ExtentX        =   2196
      _ExtentY        =   2178
      _StockProps     =   66
      Caption         =   "P/R"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      Shape           =   6
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":104EB
      textLT          =   "frmSales1.frx":10551
      textCT          =   "frmSales1.frx":10569
      textRT          =   "frmSales1.frx":10581
      textLM          =   "frmSales1.frx":10599
      textRM          =   "frmSales1.frx":105B1
      textLB          =   "frmSales1.frx":105C9
      textCB          =   "frmSales1.frx":105E1
      textRB          =   "frmSales1.frx":105F9
      colorBack       =   "frmSales1.frx":10611
      colorIntern     =   "frmSales1.frx":1063B
      colorMO         =   "frmSales1.frx":10665
      colorFocus      =   "frmSales1.frx":1068F
      colorDisabled   =   "frmSales1.frx":106B9
      colorPressed    =   "frmSales1.frx":106E3
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   885
      Left            =   5700
      TabIndex        =   33
      Top             =   345
      Visible         =   0   'False
      Width           =   8295
      _Version        =   524298
      _ExtentX        =   14631
      _ExtentY        =   1561
      _StockProps     =   66
      Caption         =   "Invalid Key Pressed"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      CornerFactor    =   15
      BackColorContainer=   12632256
      SpecialEffect   =   3
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":1070D
      textLT          =   "frmSales1.frx":10793
      textCT          =   "frmSales1.frx":107AB
      textRT          =   "frmSales1.frx":107C3
      textLM          =   "frmSales1.frx":107DB
      textRM          =   "frmSales1.frx":107F3
      textLB          =   "frmSales1.frx":1080B
      textCB          =   "frmSales1.frx":10823
      textRB          =   "frmSales1.frx":1083B
      colorBack       =   "frmSales1.frx":10853
      colorIntern     =   "frmSales1.frx":1087D
      colorMO         =   "frmSales1.frx":108A7
      colorFocus      =   "frmSales1.frx":108D1
      colorDisabled   =   "frmSales1.frx":108FB
      colorPressed    =   "frmSales1.frx":10925
   End
   Begin VB.PictureBox picDigit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   13230
      Picture         =   "frmSales1.frx":1094F
      ScaleHeight     =   795
      ScaleWidth      =   525
      TabIndex        =   34
      Top             =   420
      Visible         =   0   'False
      Width           =   525
      Begin VB.Label lblDigit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   30
         TabIndex        =   35
         Top             =   150
         Width           =   435
      End
   End
   Begin VB.PictureBox picHoldFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3EEEF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1500
      ScaleHeight     =   615
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   2640
      Width           =   825
   End
   Begin VB.Timer errTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid grdPlu 
      Height          =   5910
      Left            =   180
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   75
      _cx             =   132
      _cy             =   10425
      Appearance      =   0
      BorderStyle     =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   15790305
      BackColorAlternate=   16381166
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   1500
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VSFlex8Ctl.VSFlexGrid grdDept 
      Height          =   9330
      Left            =   0
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   75
      _cx             =   132
      _cy             =   16457
      Appearance      =   0
      BorderStyle     =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   15790305
      BackColorAlternate=   16381166
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   1500
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin btButtonEx.ButtonEx cmdDept 
      Height          =   1155
      Index           =   6
      Left            =   360
      TabIndex        =   5
      Top             =   10080
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2037
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "No Sale"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdDept 
      Height          =   1095
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   9000
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 30"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdDept 
      Height          =   1095
      Index           =   4
      Left            =   360
      TabIndex        =   7
      Top             =   7950
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 20"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdDept 
      Height          =   1095
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   6900
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 10"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdDept 
      Height          =   1095
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   5850
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 4"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdDept 
      Height          =   1155
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   4740
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2037
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 3"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4470
      Top             =   10770
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1740
      Top             =   450
   End
   Begin btButtonEx.ButtonEx cmdNext 
      Height          =   930
      Index           =   0
      Left            =   4590
      TabIndex        =   1
      Top             =   330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1640
      Appearance      =   3
      BackColor       =   2720171
      Caption         =   "7"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   26.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdNext 
      Height          =   930
      Index           =   1
      Left            =   14130
      TabIndex        =   2
      Top             =   330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1640
      Appearance      =   3
      BackColor       =   2720171
      Caption         =   "8"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   26.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdDept 
      Height          =   1185
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   3600
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2090
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 2"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   0
      Left            =   1620
      TabIndex        =   12
      Top             =   3810
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":10C88
      textLT          =   "frmSales1.frx":10CA0
      textCT          =   "frmSales1.frx":10CB8
      textRT          =   "frmSales1.frx":10CD0
      textLM          =   "frmSales1.frx":10CE8
      textRM          =   "frmSales1.frx":10D00
      textLB          =   "frmSales1.frx":10D18
      textCB          =   "frmSales1.frx":10D30
      textRB          =   "frmSales1.frx":10D48
      colorBack       =   "frmSales1.frx":10D60
      colorIntern     =   "frmSales1.frx":10D8A
      colorMO         =   "frmSales1.frx":10DB4
      colorFocus      =   "frmSales1.frx":10DDE
      colorDisabled   =   "frmSales1.frx":10E08
      colorPressed    =   "frmSales1.frx":10E32
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   3
      Left            =   1620
      TabIndex        =   13
      Top             =   4935
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":10E5C
      textLT          =   "frmSales1.frx":10E74
      textCT          =   "frmSales1.frx":10E8C
      textRT          =   "frmSales1.frx":10EA4
      textLM          =   "frmSales1.frx":10EBC
      textRM          =   "frmSales1.frx":10ED4
      textLB          =   "frmSales1.frx":10EEC
      textCB          =   "frmSales1.frx":10F04
      textRB          =   "frmSales1.frx":10F1C
      colorBack       =   "frmSales1.frx":10F34
      colorIntern     =   "frmSales1.frx":10F5E
      colorMO         =   "frmSales1.frx":10F88
      colorFocus      =   "frmSales1.frx":10FB2
      colorDisabled   =   "frmSales1.frx":10FDC
      colorPressed    =   "frmSales1.frx":11006
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   5
      Left            =   5490
      TabIndex        =   14
      Top             =   4935
      Visible         =   0   'False
      Width           =   1945
      _Version        =   524298
      _ExtentX        =   3431
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      PicturePosition =   9
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":11030
      textLT          =   "frmSales1.frx":11048
      textCT          =   "frmSales1.frx":11060
      textRT          =   "frmSales1.frx":11078
      textLM          =   "frmSales1.frx":11090
      textRM          =   "frmSales1.frx":110A8
      textLB          =   "frmSales1.frx":110C0
      textCB          =   "frmSales1.frx":110D8
      textRB          =   "frmSales1.frx":110F0
      colorBack       =   "frmSales1.frx":11108
      colorIntern     =   "frmSales1.frx":11132
      colorMO         =   "frmSales1.frx":1115C
      colorFocus      =   "frmSales1.frx":11186
      colorDisabled   =   "frmSales1.frx":111B0
      colorPressed    =   "frmSales1.frx":111DA
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   6
      Left            =   1620
      TabIndex        =   15
      Top             =   6060
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":11204
      textLT          =   "frmSales1.frx":1121C
      textCT          =   "frmSales1.frx":11234
      textRT          =   "frmSales1.frx":1124C
      textLM          =   "frmSales1.frx":11264
      textRM          =   "frmSales1.frx":1127C
      textLB          =   "frmSales1.frx":11294
      textCB          =   "frmSales1.frx":112AC
      textRB          =   "frmSales1.frx":112C4
      colorBack       =   "frmSales1.frx":112DC
      colorIntern     =   "frmSales1.frx":11306
      colorMO         =   "frmSales1.frx":11330
      colorFocus      =   "frmSales1.frx":1135A
      colorDisabled   =   "frmSales1.frx":11384
      colorPressed    =   "frmSales1.frx":113AE
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   8
      Left            =   5490
      TabIndex        =   16
      Top             =   6060
      Visible         =   0   'False
      Width           =   1945
      _Version        =   524298
      _ExtentX        =   3431
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":113D8
      textLT          =   "frmSales1.frx":113F0
      textCT          =   "frmSales1.frx":11408
      textRT          =   "frmSales1.frx":11420
      textLM          =   "frmSales1.frx":11438
      textRM          =   "frmSales1.frx":11450
      textLB          =   "frmSales1.frx":11468
      textCB          =   "frmSales1.frx":11480
      textRB          =   "frmSales1.frx":11498
      colorBack       =   "frmSales1.frx":114B0
      colorIntern     =   "frmSales1.frx":114DA
      colorMO         =   "frmSales1.frx":11504
      colorFocus      =   "frmSales1.frx":1152E
      colorDisabled   =   "frmSales1.frx":11558
      colorPressed    =   "frmSales1.frx":11582
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   9
      Left            =   1620
      TabIndex        =   17
      Top             =   7185
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":115AC
      textLT          =   "frmSales1.frx":115C4
      textCT          =   "frmSales1.frx":115DC
      textRT          =   "frmSales1.frx":115F4
      textLM          =   "frmSales1.frx":1160C
      textRM          =   "frmSales1.frx":11624
      textLB          =   "frmSales1.frx":1163C
      textCB          =   "frmSales1.frx":11654
      textRB          =   "frmSales1.frx":1166C
      colorBack       =   "frmSales1.frx":11684
      colorIntern     =   "frmSales1.frx":116AE
      colorMO         =   "frmSales1.frx":116D8
      colorFocus      =   "frmSales1.frx":11702
      colorDisabled   =   "frmSales1.frx":1172C
      colorPressed    =   "frmSales1.frx":11756
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   11
      Left            =   5490
      TabIndex        =   18
      Top             =   7185
      Visible         =   0   'False
      Width           =   1945
      _Version        =   524298
      _ExtentX        =   3431
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      PicturePosition =   9
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":11780
      textLT          =   "frmSales1.frx":11798
      textCT          =   "frmSales1.frx":117B0
      textRT          =   "frmSales1.frx":117C8
      textLM          =   "frmSales1.frx":117E0
      textRM          =   "frmSales1.frx":117F8
      textLB          =   "frmSales1.frx":11810
      textCB          =   "frmSales1.frx":11828
      textRB          =   "frmSales1.frx":11840
      colorBack       =   "frmSales1.frx":11858
      colorIntern     =   "frmSales1.frx":11882
      colorMO         =   "frmSales1.frx":118AC
      colorFocus      =   "frmSales1.frx":118D6
      colorDisabled   =   "frmSales1.frx":11900
      colorPressed    =   "frmSales1.frx":1192A
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   12
      Left            =   1620
      TabIndex        =   19
      Top             =   8310
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":11954
      textLT          =   "frmSales1.frx":1196C
      textCT          =   "frmSales1.frx":11984
      textRT          =   "frmSales1.frx":1199C
      textLM          =   "frmSales1.frx":119B4
      textRM          =   "frmSales1.frx":119CC
      textLB          =   "frmSales1.frx":119E4
      textCB          =   "frmSales1.frx":119FC
      textRB          =   "frmSales1.frx":11A14
      colorBack       =   "frmSales1.frx":11A2C
      colorIntern     =   "frmSales1.frx":11A56
      colorMO         =   "frmSales1.frx":11A80
      colorFocus      =   "frmSales1.frx":11AAA
      colorDisabled   =   "frmSales1.frx":11AD4
      colorPressed    =   "frmSales1.frx":11AFE
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   14
      Left            =   5490
      TabIndex        =   20
      Top             =   8310
      Visible         =   0   'False
      Width           =   1945
      _Version        =   524298
      _ExtentX        =   3431
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":11B28
      textLT          =   "frmSales1.frx":11B40
      textCT          =   "frmSales1.frx":11B58
      textRT          =   "frmSales1.frx":11B70
      textLM          =   "frmSales1.frx":11B88
      textRM          =   "frmSales1.frx":11BA0
      textLB          =   "frmSales1.frx":11BB8
      textCB          =   "frmSales1.frx":11BD0
      textRB          =   "frmSales1.frx":11BE8
      colorBack       =   "frmSales1.frx":11C00
      colorIntern     =   "frmSales1.frx":11C2A
      colorMO         =   "frmSales1.frx":11C54
      colorFocus      =   "frmSales1.frx":11C7E
      colorDisabled   =   "frmSales1.frx":11CA8
      colorPressed    =   "frmSales1.frx":11CD2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   15
      Left            =   1620
      TabIndex        =   21
      Top             =   9435
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1993
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":11CFC
      textLT          =   "frmSales1.frx":11D14
      textCT          =   "frmSales1.frx":11D2C
      textRT          =   "frmSales1.frx":11D44
      textLM          =   "frmSales1.frx":11D5C
      textRM          =   "frmSales1.frx":11D74
      textLB          =   "frmSales1.frx":11D8C
      textCB          =   "frmSales1.frx":11DA4
      textRB          =   "frmSales1.frx":11DBC
      colorBack       =   "frmSales1.frx":11DD4
      colorIntern     =   "frmSales1.frx":11DFE
      colorMO         =   "frmSales1.frx":11E28
      colorFocus      =   "frmSales1.frx":11E52
      colorDisabled   =   "frmSales1.frx":11E7C
      colorPressed    =   "frmSales1.frx":11EA6
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   17
      Left            =   5490
      TabIndex        =   22
      Top             =   9435
      Visible         =   0   'False
      Width           =   1945
      _Version        =   524298
      _ExtentX        =   3431
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":11ED0
      textLT          =   "frmSales1.frx":11EE8
      textCT          =   "frmSales1.frx":11F00
      textRT          =   "frmSales1.frx":11F18
      textLM          =   "frmSales1.frx":11F30
      textRM          =   "frmSales1.frx":11F48
      textLB          =   "frmSales1.frx":11F60
      textCB          =   "frmSales1.frx":11F78
      textRB          =   "frmSales1.frx":11F90
      colorBack       =   "frmSales1.frx":11FA8
      colorIntern     =   "frmSales1.frx":11FD2
      colorMO         =   "frmSales1.frx":11FFC
      colorFocus      =   "frmSales1.frx":12026
      colorDisabled   =   "frmSales1.frx":12050
      colorPressed    =   "frmSales1.frx":1207A
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   1
      Left            =   3555
      TabIndex        =   23
      Top             =   3810
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":120A4
      textLT          =   "frmSales1.frx":120BC
      textCT          =   "frmSales1.frx":120D4
      textRT          =   "frmSales1.frx":120EC
      textLM          =   "frmSales1.frx":12104
      textRM          =   "frmSales1.frx":1211C
      textLB          =   "frmSales1.frx":12134
      textCB          =   "frmSales1.frx":1214C
      textRB          =   "frmSales1.frx":12164
      colorBack       =   "frmSales1.frx":1217C
      colorIntern     =   "frmSales1.frx":121A6
      colorMO         =   "frmSales1.frx":121D0
      colorFocus      =   "frmSales1.frx":121FA
      colorDisabled   =   "frmSales1.frx":12224
      colorPressed    =   "frmSales1.frx":1224E
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   4
      Left            =   3555
      TabIndex        =   24
      Top             =   4935
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      PicturePosition =   9
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":12278
      textLT          =   "frmSales1.frx":12290
      textCT          =   "frmSales1.frx":122A8
      textRT          =   "frmSales1.frx":122C0
      textLM          =   "frmSales1.frx":122D8
      textRM          =   "frmSales1.frx":122F0
      textLB          =   "frmSales1.frx":12308
      textCB          =   "frmSales1.frx":12320
      textRB          =   "frmSales1.frx":12338
      colorBack       =   "frmSales1.frx":12350
      colorIntern     =   "frmSales1.frx":1237A
      colorMO         =   "frmSales1.frx":123A4
      colorFocus      =   "frmSales1.frx":123CE
      colorDisabled   =   "frmSales1.frx":123F8
      colorPressed    =   "frmSales1.frx":12422
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   7
      Left            =   3555
      TabIndex        =   25
      Top             =   6060
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":1244C
      textLT          =   "frmSales1.frx":12464
      textCT          =   "frmSales1.frx":1247C
      textRT          =   "frmSales1.frx":12494
      textLM          =   "frmSales1.frx":124AC
      textRM          =   "frmSales1.frx":124C4
      textLB          =   "frmSales1.frx":124DC
      textCB          =   "frmSales1.frx":124F4
      textRB          =   "frmSales1.frx":1250C
      colorBack       =   "frmSales1.frx":12524
      colorIntern     =   "frmSales1.frx":1254E
      colorMO         =   "frmSales1.frx":12578
      colorFocus      =   "frmSales1.frx":125A2
      colorDisabled   =   "frmSales1.frx":125CC
      colorPressed    =   "frmSales1.frx":125F6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   10
      Left            =   3555
      TabIndex        =   26
      Top             =   7185
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      PicturePosition =   9
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":12620
      textLT          =   "frmSales1.frx":12638
      textCT          =   "frmSales1.frx":12650
      textRT          =   "frmSales1.frx":12668
      textLM          =   "frmSales1.frx":12680
      textRM          =   "frmSales1.frx":12698
      textLB          =   "frmSales1.frx":126B0
      textCB          =   "frmSales1.frx":126C8
      textRB          =   "frmSales1.frx":126E0
      colorBack       =   "frmSales1.frx":126F8
      colorIntern     =   "frmSales1.frx":12722
      colorMO         =   "frmSales1.frx":1274C
      colorFocus      =   "frmSales1.frx":12776
      colorDisabled   =   "frmSales1.frx":127A0
      colorPressed    =   "frmSales1.frx":127CA
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   13
      Left            =   3555
      TabIndex        =   27
      Top             =   8310
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":127F4
      textLT          =   "frmSales1.frx":1280C
      textCT          =   "frmSales1.frx":12824
      textRT          =   "frmSales1.frx":1283C
      textLM          =   "frmSales1.frx":12854
      textRM          =   "frmSales1.frx":1286C
      textLB          =   "frmSales1.frx":12884
      textCB          =   "frmSales1.frx":1289C
      textRB          =   "frmSales1.frx":128B4
      colorBack       =   "frmSales1.frx":128CC
      colorIntern     =   "frmSales1.frx":128F6
      colorMO         =   "frmSales1.frx":12920
      colorFocus      =   "frmSales1.frx":1294A
      colorDisabled   =   "frmSales1.frx":12974
      colorPressed    =   "frmSales1.frx":1299E
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   16
      Left            =   3555
      TabIndex        =   28
      Top             =   9435
      Visible         =   0   'False
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":129C8
      textLT          =   "frmSales1.frx":129E0
      textCT          =   "frmSales1.frx":129F8
      textRT          =   "frmSales1.frx":12A10
      textLM          =   "frmSales1.frx":12A28
      textRM          =   "frmSales1.frx":12A40
      textLB          =   "frmSales1.frx":12A58
      textCB          =   "frmSales1.frx":12A70
      textRB          =   "frmSales1.frx":12A88
      colorBack       =   "frmSales1.frx":12AA0
      colorIntern     =   "frmSales1.frx":12ACA
      colorMO         =   "frmSales1.frx":12AF4
      colorFocus      =   "frmSales1.frx":12B1E
      colorDisabled   =   "frmSales1.frx":12B48
      colorPressed    =   "frmSales1.frx":12B72
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   8475
      Left            =   10830
      TabIndex        =   31
      Top             =   2040
      Width           =   4215
      _cx             =   7435
      _cy             =   14949
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12377839
      ForeColor       =   -2147483640
      BackColorFixed  =   7848417
      ForeColorFixed  =   -2147483630
      BackColorSel    =   9620977
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   12377839
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaper       =   "frmSales1.frx":12B9C
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   2
      Left            =   1590
      TabIndex        =   41
      Top             =   600
      Width           =   1365
      _Version        =   524298
      _ExtentX        =   2408
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Service Charge"
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
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":14D4A
      textLT          =   "frmSales1.frx":14DC6
      textCT          =   "frmSales1.frx":14DDE
      textRT          =   "frmSales1.frx":14DF6
      textLM          =   "frmSales1.frx":14E0E
      textRM          =   "frmSales1.frx":14E26
      textLB          =   "frmSales1.frx":14E3E
      textCB          =   "frmSales1.frx":14E56
      textRB          =   "frmSales1.frx":14E6E
      colorBack       =   "frmSales1.frx":14E86
      colorIntern     =   "frmSales1.frx":14EB0
      colorMO         =   "frmSales1.frx":14EDA
      colorFocus      =   "frmSales1.frx":14F04
      colorDisabled   =   "frmSales1.frx":14F2E
      colorPressed    =   "frmSales1.frx":14F58
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   3
      Left            =   2940
      TabIndex        =   42
      Top             =   600
      Width           =   1635
      _Version        =   524298
      _ExtentX        =   2884
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Place Order"
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
      Shape           =   4
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":14F82
      textLT          =   "frmSales1.frx":14FF8
      textCT          =   "frmSales1.frx":15010
      textRT          =   "frmSales1.frx":15028
      textLM          =   "frmSales1.frx":15040
      textRM          =   "frmSales1.frx":15058
      textLB          =   "frmSales1.frx":15070
      textCB          =   "frmSales1.frx":15088
      textRB          =   "frmSales1.frx":150A0
      colorBack       =   "frmSales1.frx":150B8
      colorIntern     =   "frmSales1.frx":150E2
      colorMO         =   "frmSales1.frx":1510C
      colorFocus      =   "frmSales1.frx":15136
      colorDisabled   =   "frmSales1.frx":15160
      colorPressed    =   "frmSales1.frx":1518A
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1215
      Index           =   4
      Left            =   3540
      TabIndex        =   43
      Top             =   2340
      Width           =   1965
      _Version        =   524298
      _ExtentX        =   3466
      _ExtentY        =   2143
      _StockProps     =   66
      Caption         =   "Close Table"
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
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":151B4
      textLT          =   "frmSales1.frx":1522A
      textCT          =   "frmSales1.frx":15242
      textRT          =   "frmSales1.frx":1525A
      textLM          =   "frmSales1.frx":15272
      textRM          =   "frmSales1.frx":1528A
      textLB          =   "frmSales1.frx":152A2
      textCB          =   "frmSales1.frx":152BA
      textRB          =   "frmSales1.frx":152D2
      colorBack       =   "frmSales1.frx":152EA
      colorIntern     =   "frmSales1.frx":15314
      colorMO         =   "frmSales1.frx":1533E
      colorFocus      =   "frmSales1.frx":15368
      colorDisabled   =   "frmSales1.frx":15392
      colorPressed    =   "frmSales1.frx":153BC
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   2
      Left            =   5490
      TabIndex        =   44
      Top             =   3810
      Visible         =   0   'False
      Width           =   1945
      _Version        =   524298
      _ExtentX        =   3431
      _ExtentY        =   1984
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":153E6
      textLT          =   "frmSales1.frx":153FE
      textCT          =   "frmSales1.frx":15416
      textRT          =   "frmSales1.frx":1542E
      textLM          =   "frmSales1.frx":15446
      textRM          =   "frmSales1.frx":1545E
      textLB          =   "frmSales1.frx":15476
      textCB          =   "frmSales1.frx":1548E
      textRB          =   "frmSales1.frx":154A6
      colorBack       =   "frmSales1.frx":154BE
      colorIntern     =   "frmSales1.frx":154E8
      colorMO         =   "frmSales1.frx":15512
      colorFocus      =   "frmSales1.frx":1553C
      colorDisabled   =   "frmSales1.frx":15566
      colorPressed    =   "frmSales1.frx":15590
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1215
      Index           =   5
      Left            =   5475
      TabIndex        =   45
      Top             =   2340
      Width           =   1980
      _Version        =   524298
      _ExtentX        =   3492
      _ExtentY        =   2143
      _StockProps     =   66
      Caption         =   "Kitchen Message"
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
      CornerFactor    =   19
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":155BA
      textLT          =   "frmSales1.frx":15638
      textCT          =   "frmSales1.frx":15650
      textRT          =   "frmSales1.frx":15668
      textLM          =   "frmSales1.frx":15680
      textRM          =   "frmSales1.frx":15698
      textLB          =   "frmSales1.frx":156B0
      textCB          =   "frmSales1.frx":156C8
      textRB          =   "frmSales1.frx":156E0
      colorBack       =   "frmSales1.frx":156F8
      colorIntern     =   "frmSales1.frx":15722
      colorMO         =   "frmSales1.frx":1574C
      colorFocus      =   "frmSales1.frx":15776
      colorDisabled   =   "frmSales1.frx":157A0
      colorPressed    =   "frmSales1.frx":157CA
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   945
      Index           =   6
      Left            =   4560
      TabIndex        =   46
      Top             =   1410
      Width           =   2895
      _Version        =   524298
      _ExtentX        =   5106
      _ExtentY        =   1667
      _StockProps     =   66
      Caption         =   "Print Bill"
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
      CornerFactor    =   20
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":157F4
      textLT          =   "frmSales1.frx":15868
      textCT          =   "frmSales1.frx":15880
      textRT          =   "frmSales1.frx":15898
      textLM          =   "frmSales1.frx":158B0
      textRM          =   "frmSales1.frx":158C8
      textLB          =   "frmSales1.frx":158E0
      textCB          =   "frmSales1.frx":158F8
      textRB          =   "frmSales1.frx":15910
      colorBack       =   "frmSales1.frx":15928
      colorIntern     =   "frmSales1.frx":15952
      colorMO         =   "frmSales1.frx":1597C
      colorFocus      =   "frmSales1.frx":159A6
      colorDisabled   =   "frmSales1.frx":159D0
      colorPressed    =   "frmSales1.frx":159FA
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   1
      Left            =   9105
      TabIndex        =   47
      Top             =   1410
      Width           =   1500
      _Version        =   524298
      _ExtentX        =   2646
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":15A24
      textLT          =   "frmSales1.frx":15A86
      textCT          =   "frmSales1.frx":15A9E
      textRT          =   "frmSales1.frx":15AB6
      textLM          =   "frmSales1.frx":15ACE
      textRM          =   "frmSales1.frx":15AE6
      textLB          =   "frmSales1.frx":15AFE
      textCB          =   "frmSales1.frx":15B16
      textRB          =   "frmSales1.frx":15B2E
      colorBack       =   "frmSales1.frx":15B46
      colorIntern     =   "frmSales1.frx":15B70
      colorMO         =   "frmSales1.frx":15B9A
      colorFocus      =   "frmSales1.frx":15BC4
      colorDisabled   =   "frmSales1.frx":15BEE
      colorPressed    =   "frmSales1.frx":15C18
      Style           =   2
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   3
      Left            =   9105
      TabIndex        =   48
      Top             =   2550
      Width           =   1500
      _Version        =   524298
      _ExtentX        =   2646
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":15C42
      textLT          =   "frmSales1.frx":15CA4
      textCT          =   "frmSales1.frx":15CBC
      textRT          =   "frmSales1.frx":15CD4
      textLM          =   "frmSales1.frx":15CEC
      textRM          =   "frmSales1.frx":15D04
      textLB          =   "frmSales1.frx":15D1C
      textCB          =   "frmSales1.frx":15D34
      textRB          =   "frmSales1.frx":15D4C
      colorBack       =   "frmSales1.frx":15D64
      colorIntern     =   "frmSales1.frx":15D8E
      colorMO         =   "frmSales1.frx":15DB8
      colorFocus      =   "frmSales1.frx":15DE2
      colorDisabled   =   "frmSales1.frx":15E0C
      colorPressed    =   "frmSales1.frx":15E36
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   5
      Left            =   9105
      TabIndex        =   49
      Top             =   3690
      Width           =   1500
      _Version        =   524298
      _ExtentX        =   2646
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":15E60
      textLT          =   "frmSales1.frx":15EC2
      textCT          =   "frmSales1.frx":15EDA
      textRT          =   "frmSales1.frx":15EF2
      textLM          =   "frmSales1.frx":15F0A
      textRM          =   "frmSales1.frx":15F22
      textLB          =   "frmSales1.frx":15F3A
      textCB          =   "frmSales1.frx":15F52
      textRB          =   "frmSales1.frx":15F6A
      colorBack       =   "frmSales1.frx":15F82
      colorIntern     =   "frmSales1.frx":15FAC
      colorMO         =   "frmSales1.frx":15FD6
      colorFocus      =   "frmSales1.frx":16000
      colorDisabled   =   "frmSales1.frx":1602A
      colorPressed    =   "frmSales1.frx":16054
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   7
      Left            =   9105
      TabIndex        =   50
      Top             =   4830
      Width           =   1500
      _Version        =   524298
      _ExtentX        =   2646
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":1607E
      textLT          =   "frmSales1.frx":160E0
      textCT          =   "frmSales1.frx":160F8
      textRT          =   "frmSales1.frx":16110
      textLM          =   "frmSales1.frx":16128
      textRM          =   "frmSales1.frx":16140
      textLB          =   "frmSales1.frx":16158
      textCB          =   "frmSales1.frx":16170
      textRB          =   "frmSales1.frx":16188
      colorBack       =   "frmSales1.frx":161A0
      colorIntern     =   "frmSales1.frx":161CA
      colorMO         =   "frmSales1.frx":161F4
      colorFocus      =   "frmSales1.frx":1621E
      colorDisabled   =   "frmSales1.frx":16248
      colorPressed    =   "frmSales1.frx":16272
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   9
      Left            =   9105
      TabIndex        =   51
      Top             =   5970
      Width           =   1500
      _Version        =   524298
      _ExtentX        =   2646
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":1629C
      textLT          =   "frmSales1.frx":162FE
      textCT          =   "frmSales1.frx":16316
      textRT          =   "frmSales1.frx":1632E
      textLM          =   "frmSales1.frx":16346
      textRM          =   "frmSales1.frx":1635E
      textLB          =   "frmSales1.frx":16376
      textCB          =   "frmSales1.frx":1638E
      textRB          =   "frmSales1.frx":163A6
      colorBack       =   "frmSales1.frx":163BE
      colorIntern     =   "frmSales1.frx":163E8
      colorMO         =   "frmSales1.frx":16412
      colorFocus      =   "frmSales1.frx":1643C
      colorDisabled   =   "frmSales1.frx":16466
      colorPressed    =   "frmSales1.frx":16490
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   11
      Left            =   9105
      TabIndex        =   52
      Top             =   7110
      Width           =   1500
      _Version        =   524298
      _ExtentX        =   2646
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":164BA
      textLT          =   "frmSales1.frx":1651C
      textCT          =   "frmSales1.frx":16534
      textRT          =   "frmSales1.frx":1654C
      textLM          =   "frmSales1.frx":16564
      textRM          =   "frmSales1.frx":1657C
      textLB          =   "frmSales1.frx":16594
      textCB          =   "frmSales1.frx":165AC
      textRB          =   "frmSales1.frx":165C4
      colorBack       =   "frmSales1.frx":165DC
      colorIntern     =   "frmSales1.frx":16606
      colorMO         =   "frmSales1.frx":16630
      colorFocus      =   "frmSales1.frx":1665A
      colorDisabled   =   "frmSales1.frx":16684
      colorPressed    =   "frmSales1.frx":166AE
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   13
      Left            =   9105
      TabIndex        =   53
      Top             =   8250
      Width           =   1500
      _Version        =   524298
      _ExtentX        =   2646
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":166D8
      textLT          =   "frmSales1.frx":1673A
      textCT          =   "frmSales1.frx":16752
      textRT          =   "frmSales1.frx":1676A
      textLM          =   "frmSales1.frx":16782
      textRM          =   "frmSales1.frx":1679A
      textLB          =   "frmSales1.frx":167B2
      textCB          =   "frmSales1.frx":167CA
      textRB          =   "frmSales1.frx":167E2
      colorBack       =   "frmSales1.frx":167FA
      colorIntern     =   "frmSales1.frx":16824
      colorMO         =   "frmSales1.frx":1684E
      colorFocus      =   "frmSales1.frx":16878
      colorDisabled   =   "frmSales1.frx":168A2
      colorPressed    =   "frmSales1.frx":168CC
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1170
      Index           =   15
      Left            =   9105
      TabIndex        =   54
      Top             =   9390
      Width           =   1500
      _Version        =   524298
      _ExtentX        =   2646
      _ExtentY        =   2064
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":168F6
      textLT          =   "frmSales1.frx":16958
      textCT          =   "frmSales1.frx":16970
      textRT          =   "frmSales1.frx":16988
      textLM          =   "frmSales1.frx":169A0
      textRM          =   "frmSales1.frx":169B8
      textLB          =   "frmSales1.frx":169D0
      textCB          =   "frmSales1.frx":169E8
      textRB          =   "frmSales1.frx":16A00
      colorBack       =   "frmSales1.frx":16A18
      colorIntern     =   "frmSales1.frx":16A42
      colorMO         =   "frmSales1.frx":16A6C
      colorFocus      =   "frmSales1.frx":16A96
      colorDisabled   =   "frmSales1.frx":16AC0
      colorPressed    =   "frmSales1.frx":16AEA
      Style           =   2
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   0
      Left            =   7620
      TabIndex        =   55
      Top             =   1410
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":16B14
      textLT          =   "frmSales1.frx":16B76
      textCT          =   "frmSales1.frx":16B8E
      textRT          =   "frmSales1.frx":16BA6
      textLM          =   "frmSales1.frx":16BBE
      textRM          =   "frmSales1.frx":16BD6
      textLB          =   "frmSales1.frx":16BEE
      textCB          =   "frmSales1.frx":16C06
      textRB          =   "frmSales1.frx":16C1E
      colorBack       =   "frmSales1.frx":16C36
      colorIntern     =   "frmSales1.frx":16C60
      colorMO         =   "frmSales1.frx":16C8A
      colorFocus      =   "frmSales1.frx":16CB4
      colorDisabled   =   "frmSales1.frx":16CDE
      colorPressed    =   "frmSales1.frx":16D08
      Style           =   2
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   2
      Left            =   7620
      TabIndex        =   56
      Top             =   2550
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":16D32
      textLT          =   "frmSales1.frx":16D94
      textCT          =   "frmSales1.frx":16DAC
      textRT          =   "frmSales1.frx":16DC4
      textLM          =   "frmSales1.frx":16DDC
      textRM          =   "frmSales1.frx":16DF4
      textLB          =   "frmSales1.frx":16E0C
      textCB          =   "frmSales1.frx":16E24
      textRB          =   "frmSales1.frx":16E3C
      colorBack       =   "frmSales1.frx":16E54
      colorIntern     =   "frmSales1.frx":16E7E
      colorMO         =   "frmSales1.frx":16EA8
      colorFocus      =   "frmSales1.frx":16ED2
      colorDisabled   =   "frmSales1.frx":16EFC
      colorPressed    =   "frmSales1.frx":16F26
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   4
      Left            =   7620
      TabIndex        =   57
      Top             =   3690
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":16F50
      textLT          =   "frmSales1.frx":16FB2
      textCT          =   "frmSales1.frx":16FCA
      textRT          =   "frmSales1.frx":16FE2
      textLM          =   "frmSales1.frx":16FFA
      textRM          =   "frmSales1.frx":17012
      textLB          =   "frmSales1.frx":1702A
      textCB          =   "frmSales1.frx":17042
      textRB          =   "frmSales1.frx":1705A
      colorBack       =   "frmSales1.frx":17072
      colorIntern     =   "frmSales1.frx":1709C
      colorMO         =   "frmSales1.frx":170C6
      colorFocus      =   "frmSales1.frx":170F0
      colorDisabled   =   "frmSales1.frx":1711A
      colorPressed    =   "frmSales1.frx":17144
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   6
      Left            =   7620
      TabIndex        =   58
      Top             =   4830
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":1716E
      textLT          =   "frmSales1.frx":171D0
      textCT          =   "frmSales1.frx":171E8
      textRT          =   "frmSales1.frx":17200
      textLM          =   "frmSales1.frx":17218
      textRM          =   "frmSales1.frx":17230
      textLB          =   "frmSales1.frx":17248
      textCB          =   "frmSales1.frx":17260
      textRB          =   "frmSales1.frx":17278
      colorBack       =   "frmSales1.frx":17290
      colorIntern     =   "frmSales1.frx":172BA
      colorMO         =   "frmSales1.frx":172E4
      colorFocus      =   "frmSales1.frx":1730E
      colorDisabled   =   "frmSales1.frx":17338
      colorPressed    =   "frmSales1.frx":17362
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   8
      Left            =   7620
      TabIndex        =   59
      Top             =   5970
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":1738C
      textLT          =   "frmSales1.frx":173EE
      textCT          =   "frmSales1.frx":17406
      textRT          =   "frmSales1.frx":1741E
      textLM          =   "frmSales1.frx":17436
      textRM          =   "frmSales1.frx":1744E
      textLB          =   "frmSales1.frx":17466
      textCB          =   "frmSales1.frx":1747E
      textRB          =   "frmSales1.frx":17496
      colorBack       =   "frmSales1.frx":174AE
      colorIntern     =   "frmSales1.frx":174D8
      colorMO         =   "frmSales1.frx":17502
      colorFocus      =   "frmSales1.frx":1752C
      colorDisabled   =   "frmSales1.frx":17556
      colorPressed    =   "frmSales1.frx":17580
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   10
      Left            =   7620
      TabIndex        =   60
      Top             =   7110
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":175AA
      textLT          =   "frmSales1.frx":1760C
      textCT          =   "frmSales1.frx":17624
      textRT          =   "frmSales1.frx":1763C
      textLM          =   "frmSales1.frx":17654
      textRM          =   "frmSales1.frx":1766C
      textLB          =   "frmSales1.frx":17684
      textCB          =   "frmSales1.frx":1769C
      textRB          =   "frmSales1.frx":176B4
      colorBack       =   "frmSales1.frx":176CC
      colorIntern     =   "frmSales1.frx":176F6
      colorMO         =   "frmSales1.frx":17720
      colorFocus      =   "frmSales1.frx":1774A
      colorDisabled   =   "frmSales1.frx":17774
      colorPressed    =   "frmSales1.frx":1779E
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   12
      Left            =   7620
      TabIndex        =   61
      Top             =   8250
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":177C8
      textLT          =   "frmSales1.frx":1782A
      textCT          =   "frmSales1.frx":17842
      textRT          =   "frmSales1.frx":1785A
      textLM          =   "frmSales1.frx":17872
      textRM          =   "frmSales1.frx":1788A
      textLB          =   "frmSales1.frx":178A2
      textCB          =   "frmSales1.frx":178BA
      textRB          =   "frmSales1.frx":178D2
      colorBack       =   "frmSales1.frx":178EA
      colorIntern     =   "frmSales1.frx":17914
      colorMO         =   "frmSales1.frx":1793E
      colorFocus      =   "frmSales1.frx":17968
      colorDisabled   =   "frmSales1.frx":17992
      colorPressed    =   "frmSales1.frx":179BC
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1170
      Index           =   14
      Left            =   7620
      TabIndex        =   62
      Top             =   9390
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2064
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":179E6
      textLT          =   "frmSales1.frx":17A48
      textCT          =   "frmSales1.frx":17A60
      textRT          =   "frmSales1.frx":17A78
      textLM          =   "frmSales1.frx":17A90
      textRM          =   "frmSales1.frx":17AA8
      textLB          =   "frmSales1.frx":17AC0
      textCB          =   "frmSales1.frx":17AD8
      textRB          =   "frmSales1.frx":17AF0
      colorBack       =   "frmSales1.frx":17B08
      colorIntern     =   "frmSales1.frx":17B32
      colorMO         =   "frmSales1.frx":17B5C
      colorFocus      =   "frmSales1.frx":17B86
      colorDisabled   =   "frmSales1.frx":17BB0
      colorPressed    =   "frmSales1.frx":17BDA
      Style           =   2
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMenu 
      Height          =   5385
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   105
      _cx             =   185
      _cy             =   9499
      Appearance      =   1
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1095
      Index           =   8
      Left            =   360
      TabIndex        =   65
      Top             =   2550
      Width           =   1095
      _Version        =   524298
      _ExtentX        =   1931
      _ExtentY        =   1931
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.75
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
      Shape           =   6
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   10736617
      CaptionWordWrapPerc=   80
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales1.frx":17C04
      textLT          =   "frmSales1.frx":17C1C
      textCT          =   "frmSales1.frx":17C34
      textRT          =   "frmSales1.frx":17C4C
      textLM          =   "frmSales1.frx":17C64
      textRM          =   "frmSales1.frx":17C7C
      textLB          =   "frmSales1.frx":17C94
      textCB          =   "frmSales1.frx":17CAC
      textRB          =   "frmSales1.frx":17CC4
      colorBack       =   "frmSales1.frx":17CDC
      colorIntern     =   "frmSales1.frx":17D06
      colorMO         =   "frmSales1.frx":17D30
      colorFocus      =   "frmSales1.frx":17D5A
      colorDisabled   =   "frmSales1.frx":17D84
      colorPressed    =   "frmSales1.frx":17DAE
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin MSForms.Label lblHappy1 
      Height          =   315
      Left            =   1470
      TabIndex        =   67
      Top             =   240
      Visible         =   0   'False
      Width           =   2265
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Happy Hour Active"
      Size            =   "3995;556"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblHappy 
      Height          =   315
      Left            =   1470
      TabIndex        =   66
      Top             =   240
      Visible         =   0   'False
      Width           =   2265
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Happy Hour Active"
      Size            =   "3995;556"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblTable 
      Height          =   285
      Left            =   6390
      TabIndex        =   64
      Top             =   10860
      Width           =   3915
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "6906;503"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label lblCash 
      Height          =   435
      Left            =   10980
      TabIndex        =   37
      Top             =   1530
      Width           =   2115
      ForeColor       =   7555868
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "3731;767"
      FontName        =   "Calibri"
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblTender 
      Height          =   435
      Left            =   12990
      TabIndex        =   36
      Top             =   1530
      Width           =   1875
      ForeColor       =   7555868
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "3307;767"
      FontName        =   "Calibri"
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblKeyRegister 
      Height          =   645
      Left            =   5790
      TabIndex        =   32
      Top             =   510
      Width           =   7935
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "13996;1138"
      FontName        =   "Calibri"
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape shpLive 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   0
      Left            =   660
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin VB.Shape shpLive 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   1
      Left            =   915
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin VB.Shape shpLive 
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   2
      Left            =   1170
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin MSForms.Label lblUser 
      Height          =   285
      Left            =   10500
      TabIndex        =   30
      Top             =   10860
      Width           =   4365
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7699;503"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   1740
      TabIndex        =   29
      Top             =   10860
      Width           =   4635
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "8176;503"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Image Image1 
      Height          =   525
      Left            =   10830
      Top             =   1470
      Width           =   4215
      BorderColor     =   1471145
      BackColor       =   11983853
      BorderStyle     =   0
      SpecialEffect   =   6
      Size            =   "7435;926"
   End
   Begin MSForms.Image newBack 
      Height          =   1815
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "2725;3201"
   End
   Begin VB.Label lblPayReason 
      Height          =   255
      Left            =   6030
      TabIndex        =   68
      Top             =   660
      Visible         =   0   'False
      Width           =   2445
   End
End
Attribute VB_Name = "frmSales1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDept_Click(Index As Integer)
    send_data_steam_keylog (Me.Name & " - " & cmdDept(Index).Caption)
    If cmdErr.Visible = True Then Exit Sub
    Key_Function cmdDept(Index).Caption
    picHoldFocus.SetFocus
End Sub
Private Sub cmdDeptStrip_Click(Index As Integer)
    DoEvents
    send_data_steam_keylog (Me.Name & " - " & cmdDeptStrip(Index).Caption)
    If cmdDeptStrip(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        cmdDeptStrip(Index).Value = 0
        DoEvents
        grdDept.Row = grdDept.Row + 1
        For i = 0 To 15
            If grdDept.TextMatrix(grdDept.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdDeptStrip(i).Caption = ""
                    cmdDeptStrip(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                    cmdDeptStrip(i).BackColor = &H1F8AB8
                Else
                    cmdDeptStrip(i).Caption = ""
                    cmdDeptStrip(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                    cmdDeptStrip(i).BackColor = &H1F8AB8
                    grdDept.Row = grdDept.Row - 1
                    Exit For
                End If
            Else
                cmdDeptStrip(i).Caption = grdDept.TextMatrix(grdDept.Row, 0)
                cmdDeptStrip(i).Tag = grdDept.TextMatrix(grdDept.Row, 1)
                cmdDeptStrip(i).BackColor = grdDept.TextMatrix(grdDept.Row, 3)
                If grdDept.TextMatrix(grdDept.Row, 4) = "1" Then cmdDeptStrip(i).Value = 1
            End If
            If grdDept.Row = grdDept.Rows - 1 Then Exit For
            grdDept.Row = grdDept.Row + 1
        Next i
        For b = i + 1 To cmdDeptStrip.Count - 1
            cmdDeptStrip(b).Caption = "1"
            cmdDeptStrip(b).Tag = ""
            cmdDeptStrip(b).Visible = False
        Next b
    End If
    If cmdDeptStrip(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdDeptStrip(Index).Value = 0
        DoEvents
        cmdDeptStrip(0).Picture = ""
        While grdDept.TextMatrix(grdDept.Row, 1) <> "ßArrow"
            grdDept.Row = grdDept.Row - 1
        Wend
        grdDept.Row = grdDept.Row - 15
        For i = 0 To 15
            If grdDept.TextMatrix(grdDept.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdDeptStrip(i).Caption = ""
                    cmdDeptStrip(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                    cmdDeptStrip(i).BackColor = &H1F8AB8
                Else
                    cmdDeptStrip(i).Caption = ""
                    cmdDeptStrip(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                    cmdDeptStrip(i).BackColor = &H1F8AB8
                    grdDept.Row = grdDept.Row - 1
                    Exit For
                End If
            Else
                cmdDeptStrip(i).Caption = grdDept.TextMatrix(grdDept.Row, 0)
                cmdDeptStrip(i).Tag = grdDept.TextMatrix(grdDept.Row, 1)
                cmdDeptStrip(i).BackColor = grdDept.TextMatrix(grdDept.Row, 3)
                If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                If grdDept.TextMatrix(grdDept.Row, 4) = "1" Then cmdDeptStrip(i).Value = 1
            End If
            If grdDept.Row = grdDept.Rows - 1 Then Exit For
            grdDept.Row = grdDept.Row + 1
        Next i
        For b = i + 1 To cmdDeptStrip.Count - 1
            cmdDeptStrip(b).Caption = "1"
            cmdDeptStrip(b).Tag = ""
            cmdDeptStrip(b).Visible = False
        Next b
    End If
End Sub
Private Sub LoadPlu(Dept_No)
    grdPlu.Rows = 0
    cmdPlu(0).Caption = ""
    cmdPlu(0).Picture = ""
    cmdPlu(0).ToolTipText = ""
    ActiveReadServer "SELECT  Short_Description, Product_Code FROM  Products where Department_No= '" & Dept_No & " ' and Touch_Item=1 and Sales_Item=1 order by Short_Description"
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdPlu.Rows = grdPlu.Rows + 1
        If i < 17 And Not rs.EOF Then
            cmdPlu(i).Caption = Replace(rs.Fields("Short_Description"), "&", "&&")
            cmdPlu(i).Tag = rs.Fields("Product_Code")
            cmdPlu(i).ToolTipText = " Product Code: " & rs.Fields("Product_Code") & " "
            If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
            grdPlu.Row = grdPlu.Rows - 1
            grdPlu.TextMatrix(grdPlu.Rows - 1, 0) = Replace(rs.Fields("Short_Description"), "&", "&&")
            grdPlu.TextMatrix(grdPlu.Rows - 1, 1) = rs.Fields("Product_Code")
        Else
            If b = 0 Then
                grdPlu.TextMatrix(grdPlu.Rows - 1, 1) = "ßArrow"
                grdPlu.Rows = grdPlu.Rows + 1
                If i = 17 Then
                    cmdPlu(17).Caption = ""
                    cmdPlu(17).ToolTipText = ""
                    cmdPlu(17).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdPlu(17).Visible = False Then cmdPlu(17).Visible = True
                End If
            End If
            b = b + 1
            grdPlu.TextMatrix(grdPlu.Rows - 1, 0) = Replace(rs.Fields("Short_Description"), "&", "&&")
            grdPlu.TextMatrix(grdPlu.Rows - 1, 1) = rs.Fields("Product_Code")
            If b = 16 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdPlu.Count - 1
       cmdPlu(b).Caption = "1"
       cmdPlu(b).Tag = ""
       cmdPlu(b).ToolTipText = ""
       If cmdPlu(b).Visible = True Then cmdPlu(b).Visible = False
    Next b
End Sub

Private Sub cmdDeptStrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdDeptStrip(Index).Picture <> App.Path & "\icons\downArr.bmp" And cmdDeptStrip(Index).Picture <> App.Path & "\icons\upArr.bmp" Then
        LoadPlu cmdDeptStrip(Index).Tag
    End If
End Sub

Private Sub cmdDeptStrip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdDeptStrip(Index).Picture <> App.Path & "\icons\downArr.bmp" And cmdDeptStrip(Index).Picture <> App.Path & "\icons\upArr.bmp" Then
        For i = 0 To grdDept.Rows - 1
            If grdDept.TextMatrix(i, 1) = cmdDeptStrip(Index).Tag Then
                grdDept.TextMatrix(i, 4) = 1
            Else
                grdDept.TextMatrix(i, 4) = ""
            End If
        Next i
    End If
End Sub

Private Sub cmdErr_Click()
    send_data_steam_keylog (Me.Name & " - " & cmdErr.Caption)
    cmdErr.Visible = False
    cmdErr.BackColor = &HF2&
    errTimer.Enabled = False
    KeyRegister = ""
    lblKeyRegister = ""
    picDigit.Visible = False
End Sub
Private Sub cmdFancy_Click(Index As Integer)
    On Error Resume Next
    send_data_steam_keylog (Me.Name & " - " & cmdFancy(Index).Caption)
    If cmdFancy(Index).Caption = "CL" Then
        cmdErr.Visible = False
        cmdErr.BackColor = &HF2&
        errTimer.Enabled = False
    End If
    If cmdErr.Visible = True Then Exit Sub
    Key_Function cmdFancy(Index).Caption
    If cmdFancy(Index).Caption <> "Close Table" Then
        If cmdFancy(Index).Caption <> "Place Order" Then
            If cmdFancy(Index).Caption <> "View Tables" Then
                picHoldFocus.SetFocus
            End If
        End If
    End If
    On Error GoTo 0
End Sub
Private Sub cmdLogOff_Click()
    If ImPrinting = True Then Exit Sub
    send_data_steam_keylog (Me.Name & " - " & cmdLogoff.Caption)
    frmSplash.cmdChange.Visible = False
    If cmdErr.Visible = True Then
        picHoldFocus.SetFocus
        Exit Sub
    End If
    If TillData.DocNo <> 0 Then
        cmdErr.Caption = "Finalize the Sale before Logging Off"
        cmdErr.Visible = True
        errTimer.Enabled = True
        picHoldFocus.SetFocus
        Exit Sub
    End If
    If TillData.TableNo <> 0 Then
        cmdErr.Caption = "Close the Table before Logging Off"
        cmdErr.Visible = True
        errTimer.Enabled = True
        picHoldFocus.SetFocus
        Exit Sub
    End If
    frmSales1.KeyPreview = False
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
    Select Case UserRecord.uType
        Case 2
            If frmSales1.Height < 10000 Then
                frmSplash.Show
                KeyCode = 0
                Me.Hide
            Else
                Screen.MousePointer = 11
                frmMain.cmdBar(6).Enabled = True
                frmMain.Show
                Select Case frmSales.Tag
                    Case "Reservations": frmMain.cmdBar(0).Enabled = False
                    Case "Rooms": frmMain.cmdBar(1).Enabled = False
                    Case "Guests": frmMain.cmdBar(2).Enabled = False
                    Case "Checkin": frmMain.cmdBar(3).Enabled = False
                    Case "Stock": frmMain.cmdBar(7).Enabled = False
                    Case "Users": frmMain.cmdBar(4).Enabled = False
                    Case "Reports": frmMain.cmdBar(5).Enabled = False
                End Select
                If frmSales.Tag = "" Then frmDetails.Show
                frmSales.Tag = ""
                For Each Form In Forms
                    Unload frmSales
                    If Form.Name = "frmSales1" Then Unload frmSales1
                    If Form.Name = "frmRestRes" Then Unload frmRestRes
                    If Form.Name = "frmTillReport" Then Unload frmTillReport
                    If Form.Name = "frmBar" Then Unload frmBar
                Next Form
                KeyCode = 0
                Screen.MousePointer = 0
            End If
        Case 3, 4, 8
            frmSplash.Show
            KeyCode = 0
            For Each Form In Forms
                    If Form.Name = "frmSales" Then Unload frmSales
                    If Form.Name = "frmRestRes" Then Unload frmRestRes
                    If Form.Name = "frmTillReport" Then Unload frmTillReport
                    If Form.Name = "frmBar" Then Unload frmBar
                    Next Form
            Unload Me
        Case Else
            Screen.MousePointer = 11
            frmMain.cmdBar(6).Enabled = True
            frmMain.Show
            Select Case frmSales.Tag
                Case "Reservations": frmMain.cmdBar(0).Enabled = False
                Case "Rooms": frmMain.cmdBar(1).Enabled = False
                Case "Guests": frmMain.cmdBar(2).Enabled = False
                Case "Checkin": frmMain.cmdBar(3).Enabled = False
                Case "Stock": frmMain.cmdBar(7).Enabled = False
                Case "Users": frmMain.cmdBar(4).Enabled = False
                Case "Reports": frmMain.cmdBar(5).Enabled = False
            End Select
            If frmSales.Tag = "" Then frmDetails.Show
            frmSales.Tag = ""
            For Each Form In Forms
                Unload frmSales
                If Form.Name = "frmSales1" Then Unload frmSales1
                If Form.Name = "frmRestRes" Then Unload frmRestRes
                If Form.Name = "frmTillReport" Then Unload frmTillReport
                If Form.Name = "frmBar" Then Unload frmBar
            Next Form
            KeyCode = 0
            Screen.MousePointer = 0
    End Select
End Sub
Private Sub cmdNext_Click(Index As Integer)
    On Error Resume Next
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    Select Case Index
        Case 0
            Select Case UserRecord.uType
                Case 0 'Manager
                Case 1 'Night Manager
                Case 2 'Reservations Clerk
                Case 3 'Waiter
                    
                Case 4 'Barman
                Case 5 'GRV Clerck
                Case 6 'Buyer
                Case 7 'Supervisor
                Case 8 'Cashier
                Case 9 'Owner
            End Select
            If cmdErr.Visible = True Then Exit Sub
            Screen.MousePointer = 11
            picHoldFocus.SetFocus
            frmSales.Show
            DoEvents
            Me.Hide
            Screen.MousePointer = 0
            With frmSales
                .grdMain.Rows = grdMain.Rows
                .grdMain.ColHidden(14) = grdMain.ColHidden(14)
                For i = 1 To grdMain.Rows - 1
                    For b = 0 To grdMain.Cols - 1
                        .grdMain.TextMatrix(i, b) = grdMain.TextMatrix(i, b)
                    Next b
                    .grdMain.Cell(flexcpBackColor, i, 0, i, 2) = grdMain.Cell(flexcpBackColor, i, 0, i, 0)
                    .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = grdMain.Cell(flexcpBackColor, i, 14, i, 14)
                    .grdMain.Cell(flexcpForeColor, i, 0, i, 2) = grdMain.Cell(flexcpForeColor, i, 0, i, 0)
                    If grdMain.Cell(flexcpFontStrikethru, i, 0, i, 0) = True Then
                        .grdMain.Cell(flexcpFontStrikethru, i, 0, i, 2) = True
                    Else
                        .grdMain.Cell(flexcpFontStrikethru, i, 0, i, 2) = False
                    End If
                    If grdMain.Cell(flexcpFontBold, i, 0, i, 0) = True Then
                        .grdMain.Cell(flexcpFontBold, i, 0, i, 2) = True
                    Else
                        .grdMain.Cell(flexcpFontBold, i, 0, i, 2) = False
                    End If
                Next i
                .lblKeyRegister.Caption = lblKeyRegister.Caption
                .grdMain.HighLight = flexHighlightWithFocus
                .lblCash.Caption = lblCash.Caption
                .lblTender = lblTender.Caption
                If TillData.DocNo <> 0 Then
                    .cmdInput(14).Caption = "Corr"
                    .grdMain.HighLight = flexHighlightAlways
                Else
                    .grdMain.HighLight = flexHighlightWithFocus
                    .cmdInput(14).Caption = "No Sale"
                End If
                .grdMain.Row = grdMain.Row
                .grdMain.ShowCell .grdMain.Row, 0
                If .grdMain.Rows = 1 Then
                    .cmdFancy(4).Caption = "Member No"
                Else
                    .cmdFancy(4).Caption = "Discount"
                End If
                If grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Row, 0) = 12648447 Then
                    .cmdFancy(4).Caption = "Member No"
                End If
            End With
        Case 1
            Select Case UserRecord.uType
                Case 0 'Manager
                Case 1 'Night Manager
                Case 2 'Reservations Clerk
                Case 3 'Waiter
                    If UserRecord.Reports = False Then
                        DisplayErr "Access Denied"
                        On Error GoTo 0
                        Exit Sub
                    End If
                Case 4 'Barman
                Case 5 'GRV Clerck
                Case 6 'Buyer
                Case 7 'Supervisor
                Case 8 'Cashier
                    If UserRecord.Reports = False Then
                        DisplayErr "Access Denied"
                        On Error GoTo 0
                        Exit Sub
                    End If
                Case 9 'Owner
            End Select
            If cmdErr.Visible = True Then Exit Sub
            Screen.MousePointer = 11
            picHoldFocus.SetFocus
            frmBar.Show
            DoEvents
            Me.Hide
            Screen.MousePointer = 0
            With frmBar
                .grdMain.Rows = grdMain.Rows
                .grdMain.ColHidden(14) = grdMain.ColHidden(14)
                For i = 1 To grdMain.Rows - 1
                    For b = 0 To grdMain.Cols - 1
                        .grdMain.TextMatrix(i, b) = grdMain.TextMatrix(i, b)
                    Next b
                    .grdMain.Cell(flexcpBackColor, i, 0, i, 2) = grdMain.Cell(flexcpBackColor, i, 0, i, 0)
                    .grdMain.Cell(flexcpBackColor, i, 14, i, 14) = grdMain.Cell(flexcpBackColor, i, 14, i, 14)
                    .grdMain.Cell(flexcpForeColor, i, 0, i, 2) = grdMain.Cell(flexcpForeColor, i, 0, i, 0)
                    If grdMain.Cell(flexcpFontStrikethru, i, 0, i, 0) = True Then
                        .grdMain.Cell(flexcpFontStrikethru, i, 0, i, 2) = True
                    Else
                        .grdMain.Cell(flexcpFontStrikethru, i, 0, i, 2) = False
                    End If
                    If grdMain.Cell(flexcpFontBold, i, 0, i, 0) = True Then
                        .grdMain.Cell(flexcpFontBold, i, 0, i, 2) = True
                    Else
                        .grdMain.Cell(flexcpFontBold, i, 0, i, 2) = False
                    End If
                Next i
                .lblKeyRegister.Caption = lblKeyRegister.Caption
                .grdMain.HighLight = flexHighlightWithFocus
                .lblCash.Caption = lblCash.Caption
                .lblTender = lblTender.Caption
                If TillData.DocNo <> 0 Then
                    .cmdInput(14).Caption = "Corr"
                    .grdMain.HighLight = flexHighlightAlways
                Else
                    .grdMain.HighLight = flexHighlightWithFocus
                End If
                .grdMain.Row = grdMain.Row
                .grdMain.TopRow = grdMain.TopRow
            End With
    End Select
    On Error GoTo 0
End Sub
Private Sub cmdPlu_Click(Index As Integer)
    DoEvents
    send_data_steam_keylog (Me.Name & " - " & cmdPlu(Index).Caption)
    If cmdPlu(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        grdPlu.Row = grdPlu.Row + 1
        For i = 0 To 17
            If grdPlu.TextMatrix(grdPlu.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdPlu(i).Caption = ""
                    cmdPlu(i).ToolTipText = ""
                    cmdPlu(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                Else
                    cmdPlu(i).Caption = ""
                    cmdPlu(i).ToolTipText = ""
                    cmdPlu(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                    grdPlu.Row = grdPlu.Row - 1
                    Exit For
                End If
            Else
                cmdPlu(i).Caption = grdPlu.TextMatrix(grdPlu.Row, 0)
                cmdPlu(i).Tag = grdPlu.TextMatrix(grdPlu.Row, 1)
                cmdPlu(i).ToolTipText = " Product Code: " & grdPlu.TextMatrix(grdPlu.Row, 1) & " "
            End If
            If grdPlu.Row = grdPlu.Rows - 1 Then Exit For
            grdPlu.Row = grdPlu.Row + 1
        Next i
        For b = i + 1 To cmdPlu.Count - 1
            cmdPlu(b).Caption = "1"
            cmdPlu(b).ToolTipText = ""
            cmdPlu(b).Tag = ""
            cmdPlu(b).Visible = False
        Next b
    End If
    If cmdPlu(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdPlu(0).Picture = ""
        While grdPlu.TextMatrix(grdPlu.Row, 1) <> "ßArrow"
            grdPlu.Row = grdPlu.Row - 1
        Wend
        grdPlu.Row = grdPlu.Row - 17
        For i = 0 To 17
            If grdPlu.TextMatrix(grdPlu.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdPlu(i).Caption = ""
                    cmdPlu(i).ToolTipText = ""
                    cmdPlu(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                Else
                    cmdPlu(i).Caption = ""
                     cmdPlu(i).ToolTipText = ""
                    cmdPlu(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                    grdPlu.Row = grdPlu.Row - 1
                    Exit For
                End If
            Else
                cmdPlu(i).Caption = grdPlu.TextMatrix(grdPlu.Row, 0)
                cmdPlu(i).Tag = grdPlu.TextMatrix(grdPlu.Row, 1)
                cmdPlu(i).ToolTipText = " Product Code: " & grdPlu.TextMatrix(grdPlu.Row, 1) & " "
                If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
            End If
            If grdPlu.Row = grdPlu.Rows - 1 Then Exit For
            grdPlu.Row = grdPlu.Row + 1
        Next i
        For b = i + 1 To cmdPlu.Count - 1
            cmdPlu(b).Caption = "1"
            cmdPlu(b).Tag = ""
            cmdPlu(b).ToolTipText = ""
            cmdPlu(b).Visible = False
        Next b
    End If
End Sub

Private Sub cmdPlu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    fastqty = ""
    If KeyRegister <> "" Then
        If InStr(KeyRegister, Chr(215)) = 0 Then
            For i = Len(KeyRegister) To 1 Step -1
                If Asc(Mid(KeyRegister, i, 1)) < 48 Or Asc(Mid(KeyRegister, i, 1)) > 57 Then
                    fastqty = ""
                    Exit For
                End If
                fastqty = Mid(KeyRegister, i, 1) & fastqty
            Next i
        End If
    End If
    If Val(fastqty) <> 0 Then
        KeyRegister = fastqty & " " & Chr(215) & " "
        lblKeyRegister.Caption = KeyRegister
    End If
    If cmdPlu(Index).Picture <> App.Path & "\icons\downArr.bmp" And cmdPlu(Index).Picture <> App.Path & "\icons\upArr.bmp" Then
        KeyRegister = KeyRegister & cmdPlu(Index).Tag
        If cmdErr.Visible = True Then Exit Sub
        If grdMain.Enabled = True And InStr(KeyRegister, "Void") <> 0 Then
            CanVoid = False
            For i = Len(KeyRegister) To 1 Step -1
                If Asc(Mid(KeyRegister, i, 1)) < 48 Or Asc(Mid(KeyRegister, i, 1)) > 57 Then Exit For
                Product_Code = Mid(KeyRegister, i, 1) & Product_Code
            Next i
            Qty = "1"
            
            If InStr(KeyRegister, Chr(215)) <> 0 Then
                Qty = ""
                For i = InStr(KeyRegister, Chr(215)) - 2 To 1 Step -1
                    If Asc(Mid(KeyRegister, i, 1)) < 46 Or Asc(Mid(KeyRegister, i, 1)) > 57 Then Exit For
                    Qty = Mid(KeyRegister, i, 1) & Qty
                Next i
            End If
            
            If InStr(KeyRegister, "Price O/V") = 0 Then
                For i = 1 To grdMain.Rows - 1
                    If Product_Code = grdMain.TextMatrix(i, 9) Then
                        If Qty = grdMain.TextMatrix(i, 0) Then
                            If grdMain.Cell(flexcpForeColor, i, 0, i, 2) <> vbRed And grdMain.Cell(flexcpBackColor, i, 0, i, 2) <> &HC0C0FF Then
                                If grdMain.TextMatrix(i, 13) = 0 Then
                                    grdMain.Row = i
                                    CanVoid = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next i
            Else
                For i = 1 To grdMain.Rows - 1
                    If Product_Code = grdMain.TextMatrix(i, 9) Then
                        If Qty = grdMain.TextMatrix(i, 0) Then
                            If grdMain.Cell(flexcpForeColor, i, 0, i, 2) <> vbRed And grdMain.Cell(flexcpBackColor, i, 0, i, 2) <> &HC0C0FF Then
                                If grdMain.TextMatrix(i, 13) = 1 Then
                                    If InStr(KeyRegister, "Price O/V") <> 0 Then
                                        For b = InStr(KeyRegister, "(Price O/V") - 2 To 1 Step -1
                                            If Asc(Mid(KeyRegister, b, 1)) < 46 Or Asc(Mid(KeyRegister, b, 1)) > 57 Then Exit For
                                            SellPrice = Mid(KeyRegister, b, 1) & SellPrice
                                        Next b
                                    End If
                                    If SellPrice = (Val(grdMain.TextMatrix(i, 2)) / Val(grdMain.TextMatrix(i, 0))) * 100 Then
                                        grdMain.Row = i
                                        CanVoid = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
            If CanVoid = False Then
                cmdErr.Caption = "Match the Code,Qty and Price of the Item you intend to Void"
                cmdErr.Visible = True
                errTimer.Enabled = True
                picHoldFocus.SetFocus
                Exit Sub
            End If
        End If
        Key_Function "Plu"
        DoEvents
        picHoldFocus.SetFocus
    End If
End Sub
Private Sub errTimer_Timer()
    Select Case cmdErr.BackColor
        Case &HF2&      'White
            cmdErr.BackColor = &HFFFF&
        Case &HFFFF&    'Yellow
            cmdErr.BackColor = &HF2&
    End Select
End Sub
Public Sub LoadOldTable(Table_No)

    Clear_TillData
    ActiveReadServer "Select * from Table_Listing where Table_No= " & Table_No & " order by Line_No"
    grdMain.Rows = 1
    grdMain.ColHidden(14) = True
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Trim(rs.Fields("Qty") & "")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Short_Desc")
        If Val(rs.Fields("Line_Total") & "") <> 0 Then
            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Line_Total"), "0.00")
        End If
        If grdMain.ValueMatrix(grdMain.Rows - 1, 0) > 0 Then grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Line_Total"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("KeyString")
        If Trim(rs.Fields("KeyString") & "") = "Subtotal" Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0FFC0
        End If
        If Trim(rs.Fields("KeyString") & "") = "" Then
            grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC00000
            grdMain.Cell(flexcpFontBold, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = True
        End If
        'Kotie
        If (grdMain.ValueMatrix(grdMain.Rows - 1, 2) <> 0) Then
            grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &H0
            grdMain.Cell(flexcpFontBold, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = True
        End If
        
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Cost"), "00.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = rs.Fields("Tax_Rate")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = rs.Fields("Tax_Type")
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = Trim(rs.Fields("Keyregister") & "")
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = Trim(rs.Fields("Extra_Function") & "")
        grdMain.TextMatrix(grdMain.Rows - 1, 17) = Trim(rs.Fields("User_Overide") & "")
        grdMain.TextMatrix(grdMain.Rows - 1, 18) = Val(rs.Fields("Discount_Amt") & "")
        grdMain.TextMatrix(grdMain.Rows - 1, 19) = Val(rs.Fields("Dicount_Value") & "")
        If rs.Fields("Extra_Function") & "" <> "" Then
            Select Case rs.Fields("Extra_Function") & ""
                Case "Void"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0C0FF
                Case "Corr"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0C0FF
                    grdMain.Cell(flexcpFontStrikethru, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = True
                Case "Return Item"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0C0FF
                Case "Wastage"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0C0FF
                Case Else
                    If rs.Fields("Extra_Function") & "" <> "" Then
                        grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = vbRed
                    End If
            End Select
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 9) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 10) = rs.Fields("Dept_No")
        grdMain.TextMatrix(grdMain.Rows - 1, 11) = rs.Fields("Kitchen1") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 12) = rs.Fields("Kitchen2") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 13) = rs.Fields("Price_Override")
        grdMain.TextMatrix(grdMain.Rows - 1, 14) = rs.Fields("Printed") & ""
        
        If rs.Fields("Printed") & "" = "P" Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 14, grdMain.Rows - 1, 14) = &HC0FFFF
            grdMain.ColHidden(14) = False
        End If
        If rs.Fields("Printed") & "" = Chr(187) Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 14, grdMain.Rows - 1, 14) = &HC0FFC0
            grdMain.ColHidden(14) = False
        End If
        lblTable.Caption = " Table No: " & Table_No & " Covers: " & rs.Fields("Covers")
        TillData.Covers = rs.Fields("Covers")
        TillData.DocNo = rs.Fields("Doc_No")
        TillData.Account_No = rs.Fields("Member_No") & ""
        TillData.Table_Name = rs.Fields("Table_name") & ""

        If rs.Fields("Short_Desc") = "Service Charge" Then
            TillData.Tipp = rs.Fields("Line_Total")
            TillData.Keystring = "Service Charge"
        End If
        rs.MoveNext
    Wend
    rs.Close
    ActiveReadServer ("Select count(*) as cnt from Print_Journal where Doc_no = " & TillData.DocNo & " and Doc_type = 'Bill Print' and Table_no = '" & TillData.TableNo & "'")
        TillData.Print_Count = rs.Fields("cnt")
    rs.Close
    grdMain.Row = grdMain.Rows - 1
    grdMain.ShowCell grdMain.Rows - 1, 0
    Sale_Total = 0
    For i = 1 To grdMain.Rows - 1
        If grdMain.TextMatrix(i, 8) <> "Corr" Then
            If grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                Sale_Total = Val(Sale_Total) + Val(grdMain.TextMatrix(i, 2))
                If Val(grdMain.TextMatrix(i, 5)) = 0 And grdMain.ValueMatrix(i, 0) <> 0 Then
                    TillData.NonTaxableSales = TillData.NonTaxableSales + Val(grdMain.TextMatrix(i, 2))
                Else
                    If grdMain.ValueMatrix(i, 0) <> 0 Then
                        If Val(grdMain.TextMatrix(i, 5)) <> 0 Then
                            TillData.TaxableSales = TillData.TaxableSales + Val(grdMain.TextMatrix(i, 2))
                            TillData.CollectedTax = TillData.CollectedTax + (Val(grdMain.TextMatrix(i, 2))) - ((Val(grdMain.TextMatrix(i, 2))) / ((100 + Val(grdMain.TextMatrix(i, 5))) / 100))
                        Else
                            TillData.NonTaxableSales = TillData.NonTaxableSales + Val(grdMain.TextMatrix(i, 2))
                        End If
                    End If
                End If
            End If
        End If
        If grdMain.TextMatrix(i, 8) = "Corr" Then
            TillData.Corrects = TillData.Corrects + Val(grdMain.TextMatrix(i, 2))
            TillData.CorrectCount = TillData.CorrectCount + 1
        End If
        If grdMain.TextMatrix(i, 8) = "Return Item" Then
            TillData.ReturnTotal = TillData.ReturnTotal + Val(grdMain.TextMatrix(i, 2))
            TillData.ReturnCount = TillData.ReturnCount + 1
        End If
        If grdMain.TextMatrix(i, 8) = "Wastage" Then
            TillData.UllageTotal = TillData.UllageTotal + Val(grdMain.TextMatrix(i, 2))
            TillData.UllageCount = TillData.UllageCount + 1
        End If
        If grdMain.TextMatrix(i, 8) = "Void" Then
            TillData.VoidTotal = TillData.VoidTotal + Val(grdMain.TextMatrix(i, 2))
            TillData.VoidCount = TillData.VoidCount + 1
        End If
    Next i
    TillData.SaleTotal = Sale_Total
    TillData.TaxTotal = TillData.CollectedTax
    lblCash.Caption = "Subtotal"
    lblTender.Caption = Format(Sale_Total, "0.00")
    lblKeyRegister.Caption = "Table no: " & Table_No & " opened by " & UserRecord.Name
    If TillData.DocNo = 0 Then
        cmdDept(6).Caption = "No Sale"
    Else
        cmdDept(6).Caption = "Corr"
    End If
    ActiveUpdateServer "Update Table_Listing set Locked =1 where Table_No= " & Table_No
    GlobalMode = TillMode.Inputmode
    If frmInput.cmdErr.Caption <> "Select a Table to Print the Bill" Then
        Screen.MousePointer = 11
        If frmInput.cmdErr.Caption <> "Select a Table to Close" Then
            Finalizing = False
            frmSales1.Show
        End If
        frmInput.Hide
        Screen.MousePointer = 0
    End If
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    Select Case TillData.TableNo
        Case 0: cmdFancy(4).Caption = "View Tables"
        Case Else: cmdFancy(4).Caption = "Close Table"
    End Select
    If Me.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.782
            Me.Controls(i).Left = Me.Controls(i).Left * 0.776
            Me.Controls(i).Height = Me.Controls(i).Height * 0.79
            Me.Controls(i).top = Me.Controls(i).top * 0.78
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
        On Error GoTo 0
        newBack.Width = Me.Width
        newBack.Height = Me.Height
    End If
    grdMain.ColWidth(0) = grdMain.Width * 0.1
    grdMain.ColWidth(1) = grdMain.Width * 0.65
    grdMain.ColWidth(2) = grdMain.Width * 0.25
    grdMain.ColWidth(14) = 200
    frmSales1.KeyPreview = True
    Panel_no = 1
    If frmSales1.Tag = "1" Then
        frmSales1.Tag = ""
        On Error GoTo 0
        Exit Sub
    End If
    grdDept.Rows = 0
    cmdDeptStrip(0).Caption = ""
    cmdDeptStrip(0).Picture = ""
    DoEvents
    Select Case Dept_Order
        Case 0
            ActiveReadServer "Select * from Departments_Panel2 where Location_no = '" & Location_No & "' ORDER BY Dept_Parent,Dept_Name"
        Case 1
            ActiveReadServer "Select * from Departments_Panel2 ORDER BY Dept_Parent,convert(int, substring(department_no,(SELECT PATINDEX('%-%', Department_No))+1,len(department_No)))"
    End Select
    i = -1
    b = 0
    currentcolor = "&H001F8AB8"
    If rs.RecordCount > 1 Then
        Parent = rs.Fields("Dept_Parent")
    End If
    While Not rs.EOF
        i = i + 1
        grdDept.Rows = grdDept.Rows + 1
        If Parent <> rs.Fields("Dept_Parent") Then
            Select Case currentcolor
                Case "&H001F8AB8"
                    currentcolor = "&H00136291"
                Case "&H00136291"
                    currentcolor = "&H001F8AB8"
            End Select
            Parent = rs.Fields("Dept_Parent")
        End If
        If i < 15 And Not rs.EOF Then
            cmdDeptStrip(i).Caption = Replace(rs.Fields("Dept_Name"), "&", "&&")
            cmdDeptStrip(i).Tag = rs.Fields("Department_no")
            cmdDeptStrip(i).BackColor = currentcolor
            If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
            grdDept.Row = grdDept.Rows - 1
            grdDept.TextMatrix(grdDept.Rows - 1, 0) = Replace(rs.Fields("Dept_Name"), "&", "&&")
            grdDept.TextMatrix(grdDept.Rows - 1, 1) = rs.Fields("Department_No")
            grdDept.TextMatrix(grdDept.Rows - 1, 2) = rs.Fields("Dept_Parent")
            grdDept.TextMatrix(grdDept.Rows - 1, 3) = currentcolor
        Else
            If b = 0 Then
                grdDept.TextMatrix(grdDept.Rows - 1, 1) = "ßArrow"
                grdDept.Rows = grdDept.Rows + 1
                If i = 15 Then
                    cmdDeptStrip(15).Caption = ""
                    cmdDeptStrip(15).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdDeptStrip(15).Visible = False Then cmdDeptStrip(15).Visible = True
                End If
            End If
            b = b + 1
            grdDept.TextMatrix(grdDept.Rows - 1, 0) = Replace(rs.Fields("Dept_Name"), "&", "&&")
            grdDept.TextMatrix(grdDept.Rows - 1, 1) = rs.Fields("Department_No")
            grdDept.TextMatrix(grdDept.Rows - 1, 2) = rs.Fields("Dept_Parent")
            grdDept.TextMatrix(grdDept.Rows - 1, 3) = currentcolor
            If b = 14 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdDeptStrip.Count - 1
       cmdDeptStrip(b).Caption = "1"
       cmdDeptStrip(b).Tag = ""
       cmdDeptStrip(b).Visible = False
    Next b
    
    If UserRecord.uType = 3 And TillData.DocNo = 0 And TillData.TableNo = 0 Then
        frmInput.Show
        DoEvents
        On Error GoTo 0
        Exit Sub
    End If
    picHoldFocus.SetFocus
    Screen.MousePointer = 0
    If cmdPlu(0).Visible = False Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static DallasString
    If cmdErr.Visible = True Then
        cmdErr.Visible = False
        cmdErr.BackColor = &HF2&
        errTimer.Enabled = False
        KeyRegister = ""
        lblKeyRegister = ""
        picDigit.Visible = False
        Exit Sub
    End If
    Select Case KeyCode
        Case 100
            DallasString = DallasString & "ß"
        Case 38
            grdMain.SetFocus
            If grdMain.Rows > 2 And grdMain.Row = 0 Then
                grdMain.Row = 1
            End If
            grdMain.ShowCell grdMain.Row - 1, 0
        Case 40
            grdMain.SetFocus
            If grdMain.Row <> grdMain.Rows - 1 Then
                grdMain.ShowCell grdMain.Row + 1, 0
            End If
        Case 13
            If Left(DallasString, 2) = "ßß" Then
                DallasString = ""
                If grdMain.Rows > 1 Then
                    Key_Function "Place Order"
                Else
                    Key_Function "Close Table"
                End If
            Else
                Key_Function "Plu"
            End If
        Case 27
            If TillData.DocNo <> 0 Then
                cmdErr.Caption = "Finalize the Sale before Logging Off"
                cmdErr.Visible = True
                errTimer.Enabled = True
                picHoldFocus.SetFocus
                Exit Sub
            End If
            Select Case UserRecord.uType
                Case 3, 4, 8
                    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
                    frmSplash.Show
                    KeyCode = 0
                    Me.Hide
                Case Else
                    Screen.MousePointer = 11
                    frmMain.cmdBar(6).Enabled = True
                    frmMain.Show
                    frmDetails.Show
                    For Each Form In Forms
                        Unload frmSales
                        If Form.Name = "frmSales1" Then Unload frmSales1
                        If Form.Name = "frmRestRes" Then Unload frmRestRes
                        If Form.Name = "frmTillReport" Then Unload frmTillReport
                    Next Form
                    frmDetails.Show
                    KeyCode = 0
                    Screen.MousePointer = 0
            End Select
        Case 112
            Key_Function "Cash"
            KeyCode = 0
        Case 113
            Key_Function "Card"
            KeyCode = 0
        Case 114
            Key_Function "Voucher"
            KeyCode = 0
        Case 115
            Key_Function "Charge"
            KeyCode = 0
        Case 117
            Key_Function "Corr"
            KeyCode = 0
        Case 118
            Key_Function "Void"
            KeyCode = 0
        Case 119
            Key_Function "Return Item"
            KeyCode = 0
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If cmdErr.Visible = True Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case KeyAscii
        Case 8
            Key_Function "CL"
        Case 43
            Key_Function "Price O/V"
        Case 42
            Key_Function "x"
        Case 46
            Key_Function "."
        Case 48 To 57
            Key_Function (Chr(KeyAscii))
        Case 45
            Key_Function "Dept"
    End Select
    KeyAscii = 0
End Sub
Private Sub Form_Load()
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    lblKeyRegister.TextAlign = fmTextAlignLeft
    lblKeyRegister.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    grdMain.TextMatrix(0, 0) = " No"
    grdMain.TextMatrix(0, 1) = "Description"
    grdMain.TextMatrix(0, 2) = "Total "
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignRightCenter
    grdMain.ColHidden(3) = True
    grdMain.ColHidden(4) = True
    grdMain.ColHidden(5) = True
    grdMain.ColHidden(6) = True
    grdMain.ColHidden(7) = True
    grdMain.ColHidden(8) = True
    grdMain.ColHidden(9) = True
    grdMain.ColHidden(10) = True
    grdMain.ColHidden(11) = True
    grdMain.ColHidden(12) = True
    grdMain.ColHidden(13) = True
    grdMain.ColHidden(14) = True
    grdMain.ColHidden(15) = True
    grdMain.ColHidden(16) = True
    grdMain.ColHidden(17) = True
    grdMain.ColHidden(18) = True
    grdMain.ColHidden(19) = True
    grdMain.Cell(flexcpForeColor, 0, 0, 0, 2) = 0
End Sub
Private Sub grdMain_Click()
    If GlobalMode = TillMode.StartMode Or GlobalMode = TillMode.NewMode Then
        grdMain.Rows = 1
        lblTender.Caption = "0.00"
        lblCash = ""
        lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    End If
End Sub



Private Sub Timer1_Timer()
    lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
End Sub
Private Sub Timer2_Timer()
    For i = 0 To 2
        If shpLive(i).BackColor = &HFF00& Then
            shpLive(i).BackColor = &HFFFFFF
            If i = 2 Then
                    shpLive(0).BackColor = &HFF00&
                Else
                    shpLive(i + 1).BackColor = &HFF00&
            End If
            Exit For
        End If
    Next i
    If HappyHour = 1 Then
        lblHappy.Visible = True
        Select Case lblHappy.ForeColor
            Case &HFFFFFF
                lblHappy.ForeColor = vbRed
            Case Else
                lblHappy.ForeColor = &HFFFFFF
        End Select
    Else
        lblHappy.Visible = False
    End If
    
    If HappyHour1 = 1 Then
        lblHappy1.Visible = True
        Select Case lblHappy1.ForeColor
            Case &HFFFFFF
                lblHappy1.ForeColor = vbRed
            Case Else
                lblHappy1.ForeColor = &HFFFFFF
        End Select
    Else
        lblHappy1.Visible = False
    End If
End Sub

Private Sub Timer3_Timer()

            If UserRecord.uType = 3 Then
            'frmSales1.cmdDept(6).Enabled = False
            frmSales1.cmdDept(6).Caption = "Corr"
            Exit Sub
            End If
            If frmSales1.Visible = True Then
            frmSales1.cmdDept(6).Enabled = True
            End If
End Sub
