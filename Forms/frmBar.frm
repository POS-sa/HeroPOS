VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmBar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11430
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBar.frx":0000
   ScaleHeight     =   11430
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSCommLib.MSComm MSComm2 
      Left            =   2460
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3030
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdLogoff 
      Height          =   1725
      Left            =   375
      TabIndex        =   57
      Top             =   600
      Width           =   1275
      _Version        =   524298
      _ExtentX        =   2249
      _ExtentY        =   3043
      _StockProps     =   66
      Caption         =   "Log Off"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":10967
      textLT          =   "frmBar.frx":109D5
      textCT          =   "frmBar.frx":109ED
      textRT          =   "frmBar.frx":10A05
      textLM          =   "frmBar.frx":10A1D
      textRM          =   "frmBar.frx":10A35
      textLB          =   "frmBar.frx":10A4D
      textCB          =   "frmBar.frx":10A65
      textRB          =   "frmBar.frx":10A7D
      colorBack       =   "frmBar.frx":10A95
      colorIntern     =   "frmBar.frx":10ABF
      colorMO         =   "frmBar.frx":10AE9
      colorFocus      =   "frmBar.frx":10B13
      colorDisabled   =   "frmBar.frx":10B3D
      colorPressed    =   "frmBar.frx":10B67
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.Timer voidTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   555
      Index           =   8
      Left            =   360
      TabIndex        =   64
      Top             =   2550
      Width           =   555
      _Version        =   524298
      _ExtentX        =   979
      _ExtentY        =   979
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
      textCaption     =   "frmBar.frx":10B91
      textLT          =   "frmBar.frx":10BA9
      textCT          =   "frmBar.frx":10BC1
      textRT          =   "frmBar.frx":10BD9
      textLM          =   "frmBar.frx":10BF1
      textRM          =   "frmBar.frx":10C09
      textLB          =   "frmBar.frx":10C21
      textCB          =   "frmBar.frx":10C39
      textRB          =   "frmBar.frx":10C51
      colorBack       =   "frmBar.frx":10C69
      colorIntern     =   "frmBar.frx":10C93
      colorMO         =   "frmBar.frx":10CBD
      colorFocus      =   "frmBar.frx":10CE7
      colorDisabled   =   "frmBar.frx":10D11
      colorPressed    =   "frmBar.frx":10D3B
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdSlip 
      Height          =   1605
      Left            =   375
      TabIndex        =   63
      Top             =   3120
      Width           =   1050
      _Version        =   524298
      _ExtentX        =   1852
      _ExtentY        =   2831
      _StockProps     =   66
      Caption         =   "View Slip"
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
      BackColorContainer=   10736617
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":10D65
      textLT          =   "frmBar.frx":10DD7
      textCT          =   "frmBar.frx":10DEF
      textRT          =   "frmBar.frx":10E07
      textLM          =   "frmBar.frx":10E1F
      textRM          =   "frmBar.frx":10E37
      textLB          =   "frmBar.frx":10E4F
      textCB          =   "frmBar.frx":10E67
      textRB          =   "frmBar.frx":10E7F
      colorBack       =   "frmBar.frx":10E97
      colorIntern     =   "frmBar.frx":10EC1
      colorMO         =   "frmBar.frx":10EEB
      colorFocus      =   "frmBar.frx":10F15
      colorDisabled   =   "frmBar.frx":10F3F
      colorPressed    =   "frmBar.frx":10F69
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.Timer errTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4020
      Top             =   420
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   915
      Left            =   5700
      TabIndex        =   50
      Top             =   330
      Visible         =   0   'False
      Width           =   8295
      _Version        =   524298
      _ExtentX        =   14631
      _ExtentY        =   1614
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
      textCaption     =   "frmBar.frx":10F93
      textLT          =   "frmBar.frx":11019
      textCT          =   "frmBar.frx":11031
      textRT          =   "frmBar.frx":11049
      textLM          =   "frmBar.frx":11061
      textRM          =   "frmBar.frx":11079
      textLB          =   "frmBar.frx":11091
      textCB          =   "frmBar.frx":110A9
      textRB          =   "frmBar.frx":110C1
      colorBack       =   "frmBar.frx":110D9
      colorIntern     =   "frmBar.frx":11103
      colorMO         =   "frmBar.frx":1112D
      colorFocus      =   "frmBar.frx":11157
      colorDisabled   =   "frmBar.frx":11181
      colorPressed    =   "frmBar.frx":111AB
   End
   Begin VB.PictureBox picDigit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   13320
      Picture         =   "frmBar.frx":111D5
      ScaleHeight     =   795
      ScaleWidth      =   525
      TabIndex        =   51
      Top             =   390
      Visible         =   0   'False
      Width           =   525
      Begin VB.Label lblDigit 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   90
         TabIndex        =   52
         Top             =   180
         Width           =   435
      End
   End
   Begin VB.PictureBox picSlip 
      BackColor       =   &H00FFFFFF&
      Height          =   6795
      Left            =   1590
      ScaleHeight     =   6735
      ScaleWidth      =   5835
      TabIndex        =   46
      Top             =   3780
      Visible         =   0   'False
      Width           =   5895
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   600
         Index           =   1
         Left            =   5130
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   6150
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1058
         Appearance      =   3
         BackColor       =   7848417
         BorderColor     =   2130608
         Caption         =   "6"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin VSFlex8Ctl.VSFlexGrid grdMain 
         Height          =   6750
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   5115
         _cx             =   9022
         _cy             =   11906
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12.75
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
         BackColorSel    =   4827091
         ForeColorSel    =   -2147483634
         BackColorBkg    =   7453147
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
         RowHeightMin    =   550
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
         WallPaper       =   "frmBar.frx":1150E
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   555
         Index           =   0
         Left            =   5130
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   0
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   979
         Appearance      =   3
         BackColor       =   7848417
         BorderColor     =   2130608
         Caption         =   "5"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin VB.Shape picGrid 
         BackColor       =   &H00B4DAED&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0086C4E1&
         FillColor       =   &H0071B9DB&
         FillStyle       =   2  'Horizontal Line
         Height          =   5895
         Left            =   5130
         Top             =   570
         Width           =   705
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4470
      Top             =   10770
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1740
      Top             =   780
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1125
      Index           =   6
      Left            =   360
      TabIndex        =   0
      Top             =   10110
      Width           =   1085
      _ExtentX        =   1905
      _ExtentY        =   1984
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Reprint"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdSlip1 
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   9030
      Width           =   1085
      _ExtentX        =   1905
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Slip on"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1095
      Index           =   4
      Left            =   360
      TabIndex        =   2
      Top             =   7950
      Width           =   1085
      _ExtentX        =   1905
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "No Sale"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1095
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   6900
      Width           =   1085
      _ExtentX        =   1905
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 4"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1095
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   5850
      Width           =   1085
      _ExtentX        =   1905
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 3"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1155
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   4740
      Width           =   1085
      _ExtentX        =   1905
      _ExtentY        =   2037
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "X 2"
      CaptionOffsetX  =   -1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdNext 
      Height          =   940
      Index           =   0
      Left            =   4590
      TabIndex        =   6
      Top             =   330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1667
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
      Height          =   940
      Index           =   1
      Left            =   14130
      TabIndex        =   7
      Top             =   330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1667
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
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   1
      Left            =   9120
      TabIndex        =   8
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":13658
      textLT          =   "frmBar.frx":136BA
      textCT          =   "frmBar.frx":136D2
      textRT          =   "frmBar.frx":136EA
      textLM          =   "frmBar.frx":13702
      textRM          =   "frmBar.frx":1371A
      textLB          =   "frmBar.frx":13732
      textCB          =   "frmBar.frx":1374A
      textRB          =   "frmBar.frx":13762
      colorBack       =   "frmBar.frx":1377A
      colorIntern     =   "frmBar.frx":137A4
      colorMO         =   "frmBar.frx":137CE
      colorFocus      =   "frmBar.frx":137F8
      colorDisabled   =   "frmBar.frx":13822
      colorPressed    =   "frmBar.frx":1384C
      Style           =   2
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   3
      Left            =   9120
      TabIndex        =   9
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":13876
      textLT          =   "frmBar.frx":138D8
      textCT          =   "frmBar.frx":138F0
      textRT          =   "frmBar.frx":13908
      textLM          =   "frmBar.frx":13920
      textRM          =   "frmBar.frx":13938
      textLB          =   "frmBar.frx":13950
      textCB          =   "frmBar.frx":13968
      textRB          =   "frmBar.frx":13980
      colorBack       =   "frmBar.frx":13998
      colorIntern     =   "frmBar.frx":139C2
      colorMO         =   "frmBar.frx":139EC
      colorFocus      =   "frmBar.frx":13A16
      colorDisabled   =   "frmBar.frx":13A40
      colorPressed    =   "frmBar.frx":13A6A
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   5
      Left            =   9120
      TabIndex        =   10
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":13A94
      textLT          =   "frmBar.frx":13AF6
      textCT          =   "frmBar.frx":13B0E
      textRT          =   "frmBar.frx":13B26
      textLM          =   "frmBar.frx":13B3E
      textRM          =   "frmBar.frx":13B56
      textLB          =   "frmBar.frx":13B6E
      textCB          =   "frmBar.frx":13B86
      textRB          =   "frmBar.frx":13B9E
      colorBack       =   "frmBar.frx":13BB6
      colorIntern     =   "frmBar.frx":13BE0
      colorMO         =   "frmBar.frx":13C0A
      colorFocus      =   "frmBar.frx":13C34
      colorDisabled   =   "frmBar.frx":13C5E
      colorPressed    =   "frmBar.frx":13C88
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   7
      Left            =   9120
      TabIndex        =   11
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":13CB2
      textLT          =   "frmBar.frx":13D14
      textCT          =   "frmBar.frx":13D2C
      textRT          =   "frmBar.frx":13D44
      textLM          =   "frmBar.frx":13D5C
      textRM          =   "frmBar.frx":13D74
      textLB          =   "frmBar.frx":13D8C
      textCB          =   "frmBar.frx":13DA4
      textRB          =   "frmBar.frx":13DBC
      colorBack       =   "frmBar.frx":13DD4
      colorIntern     =   "frmBar.frx":13DFE
      colorMO         =   "frmBar.frx":13E28
      colorFocus      =   "frmBar.frx":13E52
      colorDisabled   =   "frmBar.frx":13E7C
      colorPressed    =   "frmBar.frx":13EA6
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   9
      Left            =   9120
      TabIndex        =   12
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":13ED0
      textLT          =   "frmBar.frx":13F32
      textCT          =   "frmBar.frx":13F4A
      textRT          =   "frmBar.frx":13F62
      textLM          =   "frmBar.frx":13F7A
      textRM          =   "frmBar.frx":13F92
      textLB          =   "frmBar.frx":13FAA
      textCB          =   "frmBar.frx":13FC2
      textRB          =   "frmBar.frx":13FDA
      colorBack       =   "frmBar.frx":13FF2
      colorIntern     =   "frmBar.frx":1401C
      colorMO         =   "frmBar.frx":14046
      colorFocus      =   "frmBar.frx":14070
      colorDisabled   =   "frmBar.frx":1409A
      colorPressed    =   "frmBar.frx":140C4
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   11
      Left            =   9120
      TabIndex        =   13
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":140EE
      textLT          =   "frmBar.frx":14150
      textCT          =   "frmBar.frx":14168
      textRT          =   "frmBar.frx":14180
      textLM          =   "frmBar.frx":14198
      textRM          =   "frmBar.frx":141B0
      textLB          =   "frmBar.frx":141C8
      textCB          =   "frmBar.frx":141E0
      textRB          =   "frmBar.frx":141F8
      colorBack       =   "frmBar.frx":14210
      colorIntern     =   "frmBar.frx":1423A
      colorMO         =   "frmBar.frx":14264
      colorFocus      =   "frmBar.frx":1428E
      colorDisabled   =   "frmBar.frx":142B8
      colorPressed    =   "frmBar.frx":142E2
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   13
      Left            =   9120
      TabIndex        =   14
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":1430C
      textLT          =   "frmBar.frx":1436E
      textCT          =   "frmBar.frx":14386
      textRT          =   "frmBar.frx":1439E
      textLM          =   "frmBar.frx":143B6
      textRM          =   "frmBar.frx":143CE
      textLB          =   "frmBar.frx":143E6
      textCB          =   "frmBar.frx":143FE
      textRB          =   "frmBar.frx":14416
      colorBack       =   "frmBar.frx":1442E
      colorIntern     =   "frmBar.frx":14458
      colorMO         =   "frmBar.frx":14482
      colorFocus      =   "frmBar.frx":144AC
      colorDisabled   =   "frmBar.frx":144D6
      colorPressed    =   "frmBar.frx":14500
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1170
      Index           =   15
      Left            =   9120
      TabIndex        =   15
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":1452A
      textLT          =   "frmBar.frx":1458C
      textCT          =   "frmBar.frx":145A4
      textRT          =   "frmBar.frx":145BC
      textLM          =   "frmBar.frx":145D4
      textRM          =   "frmBar.frx":145EC
      textLB          =   "frmBar.frx":14604
      textCB          =   "frmBar.frx":1461C
      textRB          =   "frmBar.frx":14634
      colorBack       =   "frmBar.frx":1464C
      colorIntern     =   "frmBar.frx":14676
      colorMO         =   "frmBar.frx":146A0
      colorFocus      =   "frmBar.frx":146CA
      colorDisabled   =   "frmBar.frx":146F4
      colorPressed    =   "frmBar.frx":1471E
      Style           =   2
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   0
      Left            =   7635
      TabIndex        =   16
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":14748
      textLT          =   "frmBar.frx":147AA
      textCT          =   "frmBar.frx":147C2
      textRT          =   "frmBar.frx":147DA
      textLM          =   "frmBar.frx":147F2
      textRM          =   "frmBar.frx":1480A
      textLB          =   "frmBar.frx":14822
      textCB          =   "frmBar.frx":1483A
      textRB          =   "frmBar.frx":14852
      colorBack       =   "frmBar.frx":1486A
      colorIntern     =   "frmBar.frx":14894
      colorMO         =   "frmBar.frx":148BE
      colorFocus      =   "frmBar.frx":148E8
      colorDisabled   =   "frmBar.frx":14912
      colorPressed    =   "frmBar.frx":1493C
      Style           =   2
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   2
      Left            =   7635
      TabIndex        =   17
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":14966
      textLT          =   "frmBar.frx":149C8
      textCT          =   "frmBar.frx":149E0
      textRT          =   "frmBar.frx":149F8
      textLM          =   "frmBar.frx":14A10
      textRM          =   "frmBar.frx":14A28
      textLB          =   "frmBar.frx":14A40
      textCB          =   "frmBar.frx":14A58
      textRB          =   "frmBar.frx":14A70
      colorBack       =   "frmBar.frx":14A88
      colorIntern     =   "frmBar.frx":14AB2
      colorMO         =   "frmBar.frx":14ADC
      colorFocus      =   "frmBar.frx":14B06
      colorDisabled   =   "frmBar.frx":14B30
      colorPressed    =   "frmBar.frx":14B5A
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   4
      Left            =   7635
      TabIndex        =   18
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":14B84
      textLT          =   "frmBar.frx":14BE6
      textCT          =   "frmBar.frx":14BFE
      textRT          =   "frmBar.frx":14C16
      textLM          =   "frmBar.frx":14C2E
      textRM          =   "frmBar.frx":14C46
      textLB          =   "frmBar.frx":14C5E
      textCB          =   "frmBar.frx":14C76
      textRB          =   "frmBar.frx":14C8E
      colorBack       =   "frmBar.frx":14CA6
      colorIntern     =   "frmBar.frx":14CD0
      colorMO         =   "frmBar.frx":14CFA
      colorFocus      =   "frmBar.frx":14D24
      colorDisabled   =   "frmBar.frx":14D4E
      colorPressed    =   "frmBar.frx":14D78
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   6
      Left            =   7635
      TabIndex        =   19
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":14DA2
      textLT          =   "frmBar.frx":14E04
      textCT          =   "frmBar.frx":14E1C
      textRT          =   "frmBar.frx":14E34
      textLM          =   "frmBar.frx":14E4C
      textRM          =   "frmBar.frx":14E64
      textLB          =   "frmBar.frx":14E7C
      textCB          =   "frmBar.frx":14E94
      textRB          =   "frmBar.frx":14EAC
      colorBack       =   "frmBar.frx":14EC4
      colorIntern     =   "frmBar.frx":14EEE
      colorMO         =   "frmBar.frx":14F18
      colorFocus      =   "frmBar.frx":14F42
      colorDisabled   =   "frmBar.frx":14F6C
      colorPressed    =   "frmBar.frx":14F96
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   8
      Left            =   7635
      TabIndex        =   20
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":14FC0
      textLT          =   "frmBar.frx":15022
      textCT          =   "frmBar.frx":1503A
      textRT          =   "frmBar.frx":15052
      textLM          =   "frmBar.frx":1506A
      textRM          =   "frmBar.frx":15082
      textLB          =   "frmBar.frx":1509A
      textCB          =   "frmBar.frx":150B2
      textRB          =   "frmBar.frx":150CA
      colorBack       =   "frmBar.frx":150E2
      colorIntern     =   "frmBar.frx":1510C
      colorMO         =   "frmBar.frx":15136
      colorFocus      =   "frmBar.frx":15160
      colorDisabled   =   "frmBar.frx":1518A
      colorPressed    =   "frmBar.frx":151B4
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   10
      Left            =   7635
      TabIndex        =   21
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":151DE
      textLT          =   "frmBar.frx":15240
      textCT          =   "frmBar.frx":15258
      textRT          =   "frmBar.frx":15270
      textLM          =   "frmBar.frx":15288
      textRM          =   "frmBar.frx":152A0
      textLB          =   "frmBar.frx":152B8
      textCB          =   "frmBar.frx":152D0
      textRB          =   "frmBar.frx":152E8
      colorBack       =   "frmBar.frx":15300
      colorIntern     =   "frmBar.frx":1532A
      colorMO         =   "frmBar.frx":15354
      colorFocus      =   "frmBar.frx":1537E
      colorDisabled   =   "frmBar.frx":153A8
      colorPressed    =   "frmBar.frx":153D2
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   12
      Left            =   7635
      TabIndex        =   22
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":153FC
      textLT          =   "frmBar.frx":1545E
      textCT          =   "frmBar.frx":15476
      textRT          =   "frmBar.frx":1548E
      textLM          =   "frmBar.frx":154A6
      textRM          =   "frmBar.frx":154BE
      textLB          =   "frmBar.frx":154D6
      textCB          =   "frmBar.frx":154EE
      textRB          =   "frmBar.frx":15506
      colorBack       =   "frmBar.frx":1551E
      colorIntern     =   "frmBar.frx":15548
      colorMO         =   "frmBar.frx":15572
      colorFocus      =   "frmBar.frx":1559C
      colorDisabled   =   "frmBar.frx":155C6
      colorPressed    =   "frmBar.frx":155F0
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1170
      Index           =   14
      Left            =   7635
      TabIndex        =   23
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":1561A
      textLT          =   "frmBar.frx":1567C
      textCT          =   "frmBar.frx":15694
      textRT          =   "frmBar.frx":156AC
      textLM          =   "frmBar.frx":156C4
      textRM          =   "frmBar.frx":156DC
      textLB          =   "frmBar.frx":156F4
      textCB          =   "frmBar.frx":1570C
      textRB          =   "frmBar.frx":15724
      colorBack       =   "frmBar.frx":1573C
      colorIntern     =   "frmBar.frx":15766
      colorMO         =   "frmBar.frx":15790
      colorFocus      =   "frmBar.frx":157BA
      colorDisabled   =   "frmBar.frx":157E4
      colorPressed    =   "frmBar.frx":1580E
      Style           =   2
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   0
      Left            =   1620
      TabIndex        =   25
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
      textCaption     =   "frmBar.frx":15838
      textLT          =   "frmBar.frx":15850
      textCT          =   "frmBar.frx":15868
      textRT          =   "frmBar.frx":15880
      textLM          =   "frmBar.frx":15898
      textRM          =   "frmBar.frx":158B0
      textLB          =   "frmBar.frx":158C8
      textCB          =   "frmBar.frx":158E0
      textRB          =   "frmBar.frx":158F8
      colorBack       =   "frmBar.frx":15910
      colorIntern     =   "frmBar.frx":1593A
      colorMO         =   "frmBar.frx":15964
      colorFocus      =   "frmBar.frx":1598E
      colorDisabled   =   "frmBar.frx":159B8
      colorPressed    =   "frmBar.frx":159E2
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   2
      Left            =   5490
      TabIndex        =   26
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
      textCaption     =   "frmBar.frx":15A0C
      textLT          =   "frmBar.frx":15A24
      textCT          =   "frmBar.frx":15A3C
      textRT          =   "frmBar.frx":15A54
      textLM          =   "frmBar.frx":15A6C
      textRM          =   "frmBar.frx":15A84
      textLB          =   "frmBar.frx":15A9C
      textCB          =   "frmBar.frx":15AB4
      textRB          =   "frmBar.frx":15ACC
      colorBack       =   "frmBar.frx":15AE4
      colorIntern     =   "frmBar.frx":15B0E
      colorMO         =   "frmBar.frx":15B38
      colorFocus      =   "frmBar.frx":15B62
      colorDisabled   =   "frmBar.frx":15B8C
      colorPressed    =   "frmBar.frx":15BB6
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   3
      Left            =   1620
      TabIndex        =   27
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
      textCaption     =   "frmBar.frx":15BE0
      textLT          =   "frmBar.frx":15BF8
      textCT          =   "frmBar.frx":15C10
      textRT          =   "frmBar.frx":15C28
      textLM          =   "frmBar.frx":15C40
      textRM          =   "frmBar.frx":15C58
      textLB          =   "frmBar.frx":15C70
      textCB          =   "frmBar.frx":15C88
      textRB          =   "frmBar.frx":15CA0
      colorBack       =   "frmBar.frx":15CB8
      colorIntern     =   "frmBar.frx":15CE2
      colorMO         =   "frmBar.frx":15D0C
      colorFocus      =   "frmBar.frx":15D36
      colorDisabled   =   "frmBar.frx":15D60
      colorPressed    =   "frmBar.frx":15D8A
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   5
      Left            =   5490
      TabIndex        =   28
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
      textCaption     =   "frmBar.frx":15DB4
      textLT          =   "frmBar.frx":15DCC
      textCT          =   "frmBar.frx":15DE4
      textRT          =   "frmBar.frx":15DFC
      textLM          =   "frmBar.frx":15E14
      textRM          =   "frmBar.frx":15E2C
      textLB          =   "frmBar.frx":15E44
      textCB          =   "frmBar.frx":15E5C
      textRB          =   "frmBar.frx":15E74
      colorBack       =   "frmBar.frx":15E8C
      colorIntern     =   "frmBar.frx":15EB6
      colorMO         =   "frmBar.frx":15EE0
      colorFocus      =   "frmBar.frx":15F0A
      colorDisabled   =   "frmBar.frx":15F34
      colorPressed    =   "frmBar.frx":15F5E
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   6
      Left            =   1620
      TabIndex        =   29
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
      textCaption     =   "frmBar.frx":15F88
      textLT          =   "frmBar.frx":15FA0
      textCT          =   "frmBar.frx":15FB8
      textRT          =   "frmBar.frx":15FD0
      textLM          =   "frmBar.frx":15FE8
      textRM          =   "frmBar.frx":16000
      textLB          =   "frmBar.frx":16018
      textCB          =   "frmBar.frx":16030
      textRB          =   "frmBar.frx":16048
      colorBack       =   "frmBar.frx":16060
      colorIntern     =   "frmBar.frx":1608A
      colorMO         =   "frmBar.frx":160B4
      colorFocus      =   "frmBar.frx":160DE
      colorDisabled   =   "frmBar.frx":16108
      colorPressed    =   "frmBar.frx":16132
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   8
      Left            =   5490
      TabIndex        =   30
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
      textCaption     =   "frmBar.frx":1615C
      textLT          =   "frmBar.frx":16174
      textCT          =   "frmBar.frx":1618C
      textRT          =   "frmBar.frx":161A4
      textLM          =   "frmBar.frx":161BC
      textRM          =   "frmBar.frx":161D4
      textLB          =   "frmBar.frx":161EC
      textCB          =   "frmBar.frx":16204
      textRB          =   "frmBar.frx":1621C
      colorBack       =   "frmBar.frx":16234
      colorIntern     =   "frmBar.frx":1625E
      colorMO         =   "frmBar.frx":16288
      colorFocus      =   "frmBar.frx":162B2
      colorDisabled   =   "frmBar.frx":162DC
      colorPressed    =   "frmBar.frx":16306
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   9
      Left            =   1620
      TabIndex        =   31
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
      textCaption     =   "frmBar.frx":16330
      textLT          =   "frmBar.frx":16348
      textCT          =   "frmBar.frx":16360
      textRT          =   "frmBar.frx":16378
      textLM          =   "frmBar.frx":16390
      textRM          =   "frmBar.frx":163A8
      textLB          =   "frmBar.frx":163C0
      textCB          =   "frmBar.frx":163D8
      textRB          =   "frmBar.frx":163F0
      colorBack       =   "frmBar.frx":16408
      colorIntern     =   "frmBar.frx":16432
      colorMO         =   "frmBar.frx":1645C
      colorFocus      =   "frmBar.frx":16486
      colorDisabled   =   "frmBar.frx":164B0
      colorPressed    =   "frmBar.frx":164DA
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   11
      Left            =   5490
      TabIndex        =   32
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
      textCaption     =   "frmBar.frx":16504
      textLT          =   "frmBar.frx":1651C
      textCT          =   "frmBar.frx":16534
      textRT          =   "frmBar.frx":1654C
      textLM          =   "frmBar.frx":16564
      textRM          =   "frmBar.frx":1657C
      textLB          =   "frmBar.frx":16594
      textCB          =   "frmBar.frx":165AC
      textRB          =   "frmBar.frx":165C4
      colorBack       =   "frmBar.frx":165DC
      colorIntern     =   "frmBar.frx":16606
      colorMO         =   "frmBar.frx":16630
      colorFocus      =   "frmBar.frx":1665A
      colorDisabled   =   "frmBar.frx":16684
      colorPressed    =   "frmBar.frx":166AE
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   12
      Left            =   1620
      TabIndex        =   33
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
      textCaption     =   "frmBar.frx":166D8
      textLT          =   "frmBar.frx":166F0
      textCT          =   "frmBar.frx":16708
      textRT          =   "frmBar.frx":16720
      textLM          =   "frmBar.frx":16738
      textRM          =   "frmBar.frx":16750
      textLB          =   "frmBar.frx":16768
      textCB          =   "frmBar.frx":16780
      textRB          =   "frmBar.frx":16798
      colorBack       =   "frmBar.frx":167B0
      colorIntern     =   "frmBar.frx":167DA
      colorMO         =   "frmBar.frx":16804
      colorFocus      =   "frmBar.frx":1682E
      colorDisabled   =   "frmBar.frx":16858
      colorPressed    =   "frmBar.frx":16882
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   14
      Left            =   5490
      TabIndex        =   34
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
      textCaption     =   "frmBar.frx":168AC
      textLT          =   "frmBar.frx":168C4
      textCT          =   "frmBar.frx":168DC
      textRT          =   "frmBar.frx":168F4
      textLM          =   "frmBar.frx":1690C
      textRM          =   "frmBar.frx":16924
      textLB          =   "frmBar.frx":1693C
      textCB          =   "frmBar.frx":16954
      textRB          =   "frmBar.frx":1696C
      colorBack       =   "frmBar.frx":16984
      colorIntern     =   "frmBar.frx":169AE
      colorMO         =   "frmBar.frx":169D8
      colorFocus      =   "frmBar.frx":16A02
      colorDisabled   =   "frmBar.frx":16A2C
      colorPressed    =   "frmBar.frx":16A56
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   15
      Left            =   1620
      TabIndex        =   35
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
      textCaption     =   "frmBar.frx":16A80
      textLT          =   "frmBar.frx":16A98
      textCT          =   "frmBar.frx":16AB0
      textRT          =   "frmBar.frx":16AC8
      textLM          =   "frmBar.frx":16AE0
      textRM          =   "frmBar.frx":16AF8
      textLB          =   "frmBar.frx":16B10
      textCB          =   "frmBar.frx":16B28
      textRB          =   "frmBar.frx":16B40
      colorBack       =   "frmBar.frx":16B58
      colorIntern     =   "frmBar.frx":16B82
      colorMO         =   "frmBar.frx":16BAC
      colorFocus      =   "frmBar.frx":16BD6
      colorDisabled   =   "frmBar.frx":16C00
      colorPressed    =   "frmBar.frx":16C2A
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   17
      Left            =   5490
      TabIndex        =   36
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
      textCaption     =   "frmBar.frx":16C54
      textLT          =   "frmBar.frx":16C6C
      textCT          =   "frmBar.frx":16C84
      textRT          =   "frmBar.frx":16C9C
      textLM          =   "frmBar.frx":16CB4
      textRM          =   "frmBar.frx":16CCC
      textLB          =   "frmBar.frx":16CE4
      textCB          =   "frmBar.frx":16CFC
      textRB          =   "frmBar.frx":16D14
      colorBack       =   "frmBar.frx":16D2C
      colorIntern     =   "frmBar.frx":16D56
      colorMO         =   "frmBar.frx":16D80
      colorFocus      =   "frmBar.frx":16DAA
      colorDisabled   =   "frmBar.frx":16DD4
      colorPressed    =   "frmBar.frx":16DFE
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdDept 
      Height          =   8940
      Left            =   150
      TabIndex        =   37
      Top             =   150
      Visible         =   0   'False
      Width           =   45
      _cx             =   79
      _cy             =   15769
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
   Begin VSFlex8Ctl.VSFlexGrid grdPlu 
      Height          =   5910
      Left            =   0
      TabIndex        =   38
      Top             =   0
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
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   1
      Left            =   3555
      TabIndex        =   39
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
      textCaption     =   "frmBar.frx":16E28
      textLT          =   "frmBar.frx":16E40
      textCT          =   "frmBar.frx":16E58
      textRT          =   "frmBar.frx":16E70
      textLM          =   "frmBar.frx":16E88
      textRM          =   "frmBar.frx":16EA0
      textLB          =   "frmBar.frx":16EB8
      textCB          =   "frmBar.frx":16ED0
      textRB          =   "frmBar.frx":16EE8
      colorBack       =   "frmBar.frx":16F00
      colorIntern     =   "frmBar.frx":16F2A
      colorMO         =   "frmBar.frx":16F54
      colorFocus      =   "frmBar.frx":16F7E
      colorDisabled   =   "frmBar.frx":16FA8
      colorPressed    =   "frmBar.frx":16FD2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   4
      Left            =   3555
      TabIndex        =   40
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
      textCaption     =   "frmBar.frx":16FFC
      textLT          =   "frmBar.frx":17014
      textCT          =   "frmBar.frx":1702C
      textRT          =   "frmBar.frx":17044
      textLM          =   "frmBar.frx":1705C
      textRM          =   "frmBar.frx":17074
      textLB          =   "frmBar.frx":1708C
      textCB          =   "frmBar.frx":170A4
      textRB          =   "frmBar.frx":170BC
      colorBack       =   "frmBar.frx":170D4
      colorIntern     =   "frmBar.frx":170FE
      colorMO         =   "frmBar.frx":17128
      colorFocus      =   "frmBar.frx":17152
      colorDisabled   =   "frmBar.frx":1717C
      colorPressed    =   "frmBar.frx":171A6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   7
      Left            =   3555
      TabIndex        =   41
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
      textCaption     =   "frmBar.frx":171D0
      textLT          =   "frmBar.frx":171E8
      textCT          =   "frmBar.frx":17200
      textRT          =   "frmBar.frx":17218
      textLM          =   "frmBar.frx":17230
      textRM          =   "frmBar.frx":17248
      textLB          =   "frmBar.frx":17260
      textCB          =   "frmBar.frx":17278
      textRB          =   "frmBar.frx":17290
      colorBack       =   "frmBar.frx":172A8
      colorIntern     =   "frmBar.frx":172D2
      colorMO         =   "frmBar.frx":172FC
      colorFocus      =   "frmBar.frx":17326
      colorDisabled   =   "frmBar.frx":17350
      colorPressed    =   "frmBar.frx":1737A
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   10
      Left            =   3555
      TabIndex        =   42
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
      textCaption     =   "frmBar.frx":173A4
      textLT          =   "frmBar.frx":173BC
      textCT          =   "frmBar.frx":173D4
      textRT          =   "frmBar.frx":173EC
      textLM          =   "frmBar.frx":17404
      textRM          =   "frmBar.frx":1741C
      textLB          =   "frmBar.frx":17434
      textCB          =   "frmBar.frx":1744C
      textRB          =   "frmBar.frx":17464
      colorBack       =   "frmBar.frx":1747C
      colorIntern     =   "frmBar.frx":174A6
      colorMO         =   "frmBar.frx":174D0
      colorFocus      =   "frmBar.frx":174FA
      colorDisabled   =   "frmBar.frx":17524
      colorPressed    =   "frmBar.frx":1754E
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   13
      Left            =   3555
      TabIndex        =   43
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
      textCaption     =   "frmBar.frx":17578
      textLT          =   "frmBar.frx":17590
      textCT          =   "frmBar.frx":175A8
      textRT          =   "frmBar.frx":175C0
      textLM          =   "frmBar.frx":175D8
      textRM          =   "frmBar.frx":175F0
      textLB          =   "frmBar.frx":17608
      textCB          =   "frmBar.frx":17620
      textRB          =   "frmBar.frx":17638
      colorBack       =   "frmBar.frx":17650
      colorIntern     =   "frmBar.frx":1767A
      colorMO         =   "frmBar.frx":176A4
      colorFocus      =   "frmBar.frx":176CE
      colorDisabled   =   "frmBar.frx":176F8
      colorPressed    =   "frmBar.frx":17722
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   16
      Left            =   3555
      TabIndex        =   44
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
      textCaption     =   "frmBar.frx":1774C
      textLT          =   "frmBar.frx":17764
      textCT          =   "frmBar.frx":1777C
      textRT          =   "frmBar.frx":17794
      textLM          =   "frmBar.frx":177AC
      textRM          =   "frmBar.frx":177C4
      textLB          =   "frmBar.frx":177DC
      textCB          =   "frmBar.frx":177F4
      textRB          =   "frmBar.frx":1780C
      colorBack       =   "frmBar.frx":17824
      colorIntern     =   "frmBar.frx":1784E
      colorMO         =   "frmBar.frx":17878
      colorFocus      =   "frmBar.frx":178A2
      colorDisabled   =   "frmBar.frx":178CC
      colorPressed    =   "frmBar.frx":178F6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1200
      Index           =   1
      Left            =   1620
      TabIndex        =   55
      Top             =   2340
      Width           =   1950
      _Version        =   524298
      _ExtentX        =   3440
      _ExtentY        =   2117
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
      textCaption     =   "frmBar.frx":17920
      textLT          =   "frmBar.frx":1799C
      textCT          =   "frmBar.frx":179B4
      textRT          =   "frmBar.frx":179CC
      textLM          =   "frmBar.frx":179E4
      textRM          =   "frmBar.frx":179FC
      textLB          =   "frmBar.frx":17A14
      textCB          =   "frmBar.frx":17A2C
      textRB          =   "frmBar.frx":17A44
      colorBack       =   "frmBar.frx":17A5C
      colorIntern     =   "frmBar.frx":17A86
      colorMO         =   "frmBar.frx":17AB0
      colorFocus      =   "frmBar.frx":17ADA
      colorDisabled   =   "frmBar.frx":17B04
      colorPressed    =   "frmBar.frx":17B2E
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1245
      Index           =   18
      Left            =   390
      TabIndex        =   56
      Top             =   2310
      Width           =   1245
      _Version        =   524298
      _ExtentX        =   2196
      _ExtentY        =   2196
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
      textCaption     =   "frmBar.frx":17B58
      textLT          =   "frmBar.frx":17BBE
      textCT          =   "frmBar.frx":17BD6
      textRT          =   "frmBar.frx":17BEE
      textLM          =   "frmBar.frx":17C06
      textRM          =   "frmBar.frx":17C1E
      textLB          =   "frmBar.frx":17C36
      textCB          =   "frmBar.frx":17C4E
      textRB          =   "frmBar.frx":17C66
      colorBack       =   "frmBar.frx":17C7E
      colorIntern     =   "frmBar.frx":17CA8
      colorMO         =   "frmBar.frx":17CD2
      colorFocus      =   "frmBar.frx":17CFC
      colorDisabled   =   "frmBar.frx":17D26
      colorPressed    =   "frmBar.frx":17D50
      Orientation     =   6
      TextCaptionAlignment=   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   2
      Left            =   1650
      TabIndex        =   58
      Top             =   600
      Width           =   1305
      _Version        =   524298
      _ExtentX        =   2302
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Pickup Tab"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      textCaption     =   "frmBar.frx":17D7A
      textLT          =   "frmBar.frx":17DEE
      textCT          =   "frmBar.frx":17E06
      textRT          =   "frmBar.frx":17E1E
      textLM          =   "frmBar.frx":17E36
      textRM          =   "frmBar.frx":17E4E
      textLB          =   "frmBar.frx":17E66
      textCB          =   "frmBar.frx":17E7E
      textRB          =   "frmBar.frx":17E96
      colorBack       =   "frmBar.frx":17EAE
      colorIntern     =   "frmBar.frx":17ED8
      colorMO         =   "frmBar.frx":17F02
      colorFocus      =   "frmBar.frx":17F2C
      colorDisabled   =   "frmBar.frx":17F56
      colorPressed    =   "frmBar.frx":17F80
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1205
      Index           =   4
      Left            =   3570
      TabIndex        =   60
      Top             =   2340
      Width           =   1950
      _Version        =   524298
      _ExtentX        =   3440
      _ExtentY        =   2125
      _StockProps     =   66
      Caption         =   "Close Tab"
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
      textCaption     =   "frmBar.frx":17FAA
      textLT          =   "frmBar.frx":1801C
      textCT          =   "frmBar.frx":18034
      textRT          =   "frmBar.frx":1804C
      textLM          =   "frmBar.frx":18064
      textRM          =   "frmBar.frx":1807C
      textLB          =   "frmBar.frx":18094
      textCB          =   "frmBar.frx":180AC
      textRB          =   "frmBar.frx":180C4
      colorBack       =   "frmBar.frx":180DC
      colorIntern     =   "frmBar.frx":18106
      colorMO         =   "frmBar.frx":18130
      colorFocus      =   "frmBar.frx":1815A
      colorDisabled   =   "frmBar.frx":18184
      colorPressed    =   "frmBar.frx":181AE
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1205
      Index           =   5
      Left            =   5520
      TabIndex        =   61
      Top             =   2340
      Width           =   1950
      _Version        =   524298
      _ExtentX        =   3440
      _ExtentY        =   2125
      _StockProps     =   66
      Caption         =   "Split Bill"
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
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":181D8
      textLT          =   "frmBar.frx":1824C
      textCT          =   "frmBar.frx":18264
      textRT          =   "frmBar.frx":1827C
      textLM          =   "frmBar.frx":18294
      textRM          =   "frmBar.frx":182AC
      textLB          =   "frmBar.frx":182C4
      textCB          =   "frmBar.frx":182DC
      textRB          =   "frmBar.frx":182F4
      colorBack       =   "frmBar.frx":1830C
      colorIntern     =   "frmBar.frx":18336
      colorMO         =   "frmBar.frx":18360
      colorFocus      =   "frmBar.frx":1838A
      colorDisabled   =   "frmBar.frx":183B4
      colorPressed    =   "frmBar.frx":183DE
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   945
      Index           =   6
      Left            =   4560
      TabIndex        =   62
      Top             =   1410
      Width           =   1425
      _Version        =   524298
      _ExtentX        =   2514
      _ExtentY        =   1667
      _StockProps     =   66
      Caption         =   "Print Bill"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      textCaption     =   "frmBar.frx":18408
      textLT          =   "frmBar.frx":1847C
      textCT          =   "frmBar.frx":18494
      textRT          =   "frmBar.frx":184AC
      textLM          =   "frmBar.frx":184C4
      textRM          =   "frmBar.frx":184DC
      textLB          =   "frmBar.frx":184F4
      textCB          =   "frmBar.frx":1850C
      textRB          =   "frmBar.frx":18524
      colorBack       =   "frmBar.frx":1853C
      colorIntern     =   "frmBar.frx":18566
      colorMO         =   "frmBar.frx":18590
      colorFocus      =   "frmBar.frx":185BA
      colorDisabled   =   "frmBar.frx":185E4
      colorPressed    =   "frmBar.frx":1860E
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.PictureBox picHoldFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1530
      ScaleHeight     =   615
      ScaleWidth      =   825
      TabIndex        =   54
      Top             =   2670
      Width           =   825
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMenu 
      Height          =   5385
      Left            =   0
      TabIndex        =   65
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
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   3
      Left            =   2940
      TabIndex        =   59
      Top             =   600
      Width           =   1635
      _Version        =   524298
      _ExtentX        =   2884
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Create Tab"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      CaptionWordWrapPerc=   80
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmBar.frx":18638
      textLT          =   "frmBar.frx":186AC
      textCT          =   "frmBar.frx":186C4
      textRT          =   "frmBar.frx":186DC
      textLM          =   "frmBar.frx":186F4
      textRM          =   "frmBar.frx":1870C
      textLB          =   "frmBar.frx":18724
      textCB          =   "frmBar.frx":1873C
      textRB          =   "frmBar.frx":18754
      colorBack       =   "frmBar.frx":1876C
      colorIntern     =   "frmBar.frx":18796
      colorMO         =   "frmBar.frx":187C0
      colorFocus      =   "frmBar.frx":187EA
      colorDisabled   =   "frmBar.frx":18814
      colorPressed    =   "frmBar.frx":1883E
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   0
      Left            =   10830
      TabIndex        =   67
      Top             =   6120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   1
      Left            =   12255
      TabIndex        =   68
      Top             =   6120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   2
      Left            =   13680
      TabIndex        =   69
      Top             =   6120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   3
      Left            =   10830
      TabIndex        =   70
      Top             =   7230
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   4
      Left            =   12255
      TabIndex        =   71
      Top             =   7230
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   5
      Left            =   13680
      TabIndex        =   72
      Top             =   7230
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   6
      Left            =   10830
      TabIndex        =   73
      Top             =   8340
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   7
      Left            =   12255
      TabIndex        =   74
      Top             =   8340
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   8
      Left            =   13680
      TabIndex        =   75
      Top             =   8340
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   9
      Left            =   10830
      TabIndex        =   76
      Top             =   9450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "CL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   10
      Left            =   12255
      TabIndex        =   77
      Top             =   9450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   11
      Left            =   13680
      TabIndex        =   78
      Top             =   9450
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   19
      Left            =   12960
      TabIndex        =   79
      Top             =   2700
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Charge"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   17
      Left            =   10830
      TabIndex        =   80
      Top             =   2700
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Cash"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   16
      Left            =   13680
      TabIndex        =   81
      Top             =   3840
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R50-00"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   18
      Left            =   11890
      TabIndex        =   82
      Top             =   2700
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Card"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   12
      Left            =   12255
      TabIndex        =   83
      Top             =   4980
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R10-00"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   13
      Left            =   10830
      TabIndex        =   84
      Top             =   4980
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R20-00"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   22
      Left            =   10830
      TabIndex        =   85
      Top             =   3840
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R200-00"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   21
      Left            =   12255
      TabIndex        =   86
      Top             =   3840
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R100-00"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   23
      Left            =   13680
      TabIndex        =   87
      Top             =   4980
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Subtotal"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1110
      Index           =   14
      Left            =   14020
      TabIndex        =   93
      Top             =   2700
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Price O/V"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   945
      Index           =   0
      Left            =   5970
      TabIndex        =   94
      Top             =   1410
      Width           =   1495
      _Version        =   524298
      _ExtentX        =   2637
      _ExtentY        =   1667
      _StockProps     =   66
      Caption         =   "Kitchen Message"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      textCaption     =   "frmBar.frx":18868
      textLT          =   "frmBar.frx":188E6
      textCT          =   "frmBar.frx":188FE
      textRT          =   "frmBar.frx":18916
      textLM          =   "frmBar.frx":1892E
      textRM          =   "frmBar.frx":18946
      textLB          =   "frmBar.frx":1895E
      textCB          =   "frmBar.frx":18976
      textRB          =   "frmBar.frx":1898E
      colorBack       =   "frmBar.frx":189A6
      colorIntern     =   "frmBar.frx":189D0
      colorMO         =   "frmBar.frx":189FA
      colorFocus      =   "frmBar.frx":18A24
      colorDisabled   =   "frmBar.frx":18A4E
      colorPressed    =   "frmBar.frx":18A78
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin MSForms.Label lblHappy1 
      Height          =   315
      Left            =   1440
      TabIndex        =   92
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
      Left            =   1440
      TabIndex        =   91
      Top             =   235
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
   Begin MSForms.Label lblTender 
      Height          =   615
      Left            =   10900
      TabIndex        =   90
      Top             =   1740
      Width           =   4000
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "7056;1085"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   555
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCash 
      Height          =   390
      Left            =   10905
      TabIndex        =   89
      Top             =   1455
      Width           =   3000
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "5292;688"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   88
      Top             =   270
      Width           =   2115
   End
   Begin MSForms.Label lblTab 
      Height          =   285
      Left            =   6390
      TabIndex        =   66
      Top             =   10875
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
   Begin MSForms.Label lblKeyRegister 
      Height          =   705
      Left            =   6000
      TabIndex        =   53
      Top             =   570
      Width           =   7800
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "2000"
      Size            =   "13758;1244"
      FontName        =   "Calibri"
      FontHeight      =   435
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   1800
      TabIndex        =   45
      Top             =   10875
      Width           =   5295
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "9340;503"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Shape shpLive 
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   2
      Left            =   1110
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin VB.Shape shpLive 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   1
      Left            =   855
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin VB.Shape shpLive 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   0
      Left            =   600
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin MSForms.Label lblUser 
      Height          =   285
      Left            =   10440
      TabIndex        =   24
      Top             =   10875
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
   Begin MSForms.Image newBack 
      Height          =   1815
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "2725;3201"
      Picture         =   "frmBar.frx":18AA2
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdArrow_Click(Index As Integer)
    grdMain.SetFocus
    If grdMain.Rows > 2 And grdMain.Row = 0 Then
        grdMain.Row = 1
    End If
    Select Case Index
        Case 0
            If grdMain.Row > 1 Then
                grdMain.Row = grdMain.Row - 1
            End If
            grdMain.ShowCell grdMain.Row, 0
        Case 1
            If grdMain.Row < grdMain.Rows - 1 Then
                grdMain.Row = grdMain.Row + 1
            End If
            grdMain.ShowCell grdMain.Row, 0
    End Select
End Sub
Private Sub cmdDeptStrip_Click(Index As Integer)
    If frmBar.cmdInput(19).Enabled = False Then frmBar.cmdInput(19).Enabled = True
    If Finalizing = True Then Exit Sub
    send_data_steam_keylog (Me.Name & " - Dept - " & cmdDeptStrip(Index).Caption)
    DoEvents
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
                If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                cmdDeptStrip(i).BackColor = grdDept.TextMatrix(grdDept.Row, 3)
                If grdDept.TextMatrix(grdDept.Row, 4) = "1" Then cmdDeptStrip(i).Value = 1
            End If
            If grdDept.Row = grdDept.Rows - 1 Then Exit For
            grdDept.Row = grdDept.Row + 1
        Next i
        For b = i + 1 To cmdDeptStrip.Count - 1
            cmdDeptStrip(b).Caption = "1"
            cmdDeptStrip(b).Tag = ""
            If cmdDeptStrip(b).Visible = True Then cmdDeptStrip(b).Visible = False
        Next b
    End If
End Sub

Private Sub cmdErr_Click()
    send_data_steam_keylog (Me.Name & " - Clear error - " & cmdErr.Caption)
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    cmdErr.Visible = False
    cmdErr.BackColor = &HF2&
    errTimer.Enabled = False
    KeyRegister = ""
    lblKeyRegister = ""
    picDigit.Visible = False
    On Error Resume Next
    picHoldFocus.SetFocus
    On Error GoTo 0
End Sub

Private Sub cmdFancy_Click(Index As Integer)
On Error GoTo error
    send_data_steam_keylog (Me.Name & " - " & cmdFancy(Index).Caption)
    If Finalizing = True Then Exit Sub
    If Process_Running = True Then Exit Sub 'Kotie 09/04/2013
    If cmdFancy(Index).Caption = "CL" Then
        cmdErr.Visible = False
        cmdErr.BackColor = &HF2&
        errTimer.Enabled = False
    End If
    If cmdErr.Visible = True Then Exit Sub
    Process_Running = True
    Key_Function cmdFancy(Index).Caption
    If cmdFancy(Index).Caption <> "Close Table" Then
        If cmdFancy(Index).Caption <> "Close Tab" Then
            If cmdFancy(Index).Caption <> "Place Order" Then
                If cmdFancy(Index).Caption <> "Kitchen Message" Then
                    picHoldFocus.SetFocus
                End If
            End If
        End If
    End If
    
    'Kotie 09/04/2013
error:
    Debug.Print Process_Running
    Process_Running = False
    
    
End Sub
Private Sub cmdInput_Click(Index As Integer)
    If Finalizing = True Then Exit Sub
    On Error Resume Next
    send_data_steam_keylog (Me.Name & " - " & cmdInput(Index).Caption)
    If Process_Running = True Then Exit Sub 'Kotie 09/04/2013
    If Index > 11 Then Process_Running = True 'Kotie 09/04/2013
    Select Case Index
   
        Case 0 To 8, 10, 11
            cmdInput(19).Enabled = False
        
        Case 17 To 19, 23, 9
            cmdInput(19).Enabled = True

        Case 12, 13, 16, 17, 18, 19, 21, 22, 14
            cmdInput(12).Enabled = False
            cmdInput(13).Enabled = False
            cmdInput(16).Enabled = False
            cmdInput(17).Enabled = False
            cmdInput(18).Enabled = False
            cmdInput(19).Enabled = False
            cmdInput(21).Enabled = False
            cmdInput(22).Enabled = False
            cmdInput(14).Enabled = False
            
'        Case 14
'            cmdInput(12).Enabled = False
'            cmdInput(13).Enabled = False
'            cmdInput(16).Enabled = False
'            cmdInput(17).Enabled = False
'            cmdInput(18).Enabled = False
'            cmdInput(19).Enabled = False
'            cmdInput(21).Enabled = False
'            cmdInput(22).Enabled = False
'            cmdInput(14).Enabled = False
            
'            Call frmSales.cmdInput_Click(20)
            
   
    End Select
   
    If grdMain.Enabled = False And cmdInput(Index).Caption <> "Void" Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    If cmdInput(Index).Caption = "CL" Then
        frmSales.cmdErr.Visible = False
        cmdErr.BackColor = &HF2&
        errTimer.Enabled = False
    End If
    If grdMain.Enabled = True And cmdInput(Index).Caption = "Plu" And InStr(KeyRegister, "Void") <> 0 Then
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
            On Error GoTo 0
            Exit Sub
        End If
    End If
    If grdMain.Enabled = True And cmdInput(Index).Caption = "Dept" And InStr(KeyRegister, "Void") <> 0 Then
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
        
        If InStr(KeyRegister, "Price O/V") <> 0 Then
            For i = 1 To grdMain.Rows - 1
                If Product_Code = grdMain.TextMatrix(i, 10) Then
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
            On Error GoTo 0
            Exit Sub
        End If
    End If
    If grdMain.Enabled = False And cmdInput(Index).Caption = "Void" Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
        If grdMain.TextMatrix(grdMain.Row, 9) <> "0" Then
            KeyRegister = " (Void) " & grdMain.TextMatrix(grdMain.Row, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(grdMain.Row, 2)) / Val(grdMain.TextMatrix(grdMain.Row, 0))) * 100 & " (Price O/V) " & grdMain.TextMatrix(grdMain.Row, 9)
            Key_Function "Plu"
        Else
            KeyRegister = " (Void) " & grdMain.TextMatrix(grdMain.Row, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(grdMain.Row, 2)) / Val(grdMain.TextMatrix(grdMain.Row, 0))) * 100 & " (Price O/V) " & grdMain.TextMatrix(grdMain.Row, 9)
            Key_Function "Dept"
        End If
        picHoldFocus.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    If cmdErr.Visible = True Then Exit Sub
    
    Key_Function cmdInput(Index).Caption
    picHoldFocus.SetFocus
    
    cmdInput(12).Enabled = True
    cmdInput(13).Enabled = True
    cmdInput(16).Enabled = True
    cmdInput(17).Enabled = True
    cmdInput(18).Enabled = True
    cmdInput(19).Enabled = True
    cmdInput(21).Enabled = True
    cmdInput(22).Enabled = True
    cmdInput(14).Enabled = True
    Process_Running = False 'Kotie 09/04/2013
    On Error GoTo 0
End Sub
Private Sub cmdKey_Click(Index As Integer)
send_data_steam_keylog (Me.Name & " - " & cmdKey(Index).Caption)
Thediscounttotal = 0
    If Finalizing = True Then Exit Sub
    If cmdErr.Visible = True Then Exit Sub
    Key_Function cmdKey(Index).Caption
    picHoldFocus.SetFocus
End Sub
Private Sub cmdLogOff_Click()
    send_data_steam_keylog (Me.Name & " - Log off")
    If Finalizing = True Then Exit Sub
    If ImPrinting = True Then Exit Sub
    frmSplash.cmdChange.Visible = False
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
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
    KeyPreview = False
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
    Select Case UserRecord.uType
        Case 2
            If frmBar.Height < 10000 Then
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
        Case 3, 8
        
            frmSplash.Show
            KeyCode = 0
            For Each Form In Forms
                    If Form.Name = "frmSales1" Then Unload frmSales1
                    If Form.Name = "frmRestRes" Then Unload frmRestRes
                    If Form.Name = "frmTillReport" Then Unload frmTillReport
                    If Form.Name = "frmSales" Then Unload frmSales
                    Next Form
            Unload frmBar
            'Me.Hide
        
        Case 4
        frmSplash.Show
        frmBar.Hide
        
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
send_data_steam_keylog (Me.Name & " - " & cmdNext(Index).Caption)
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    Select Case Index
        Case 0
            If cmdErr.Visible = True Then Exit Sub
            If UserRecord.uType = 4 Or UserRecord.uType = 8 Then
                If UserRecord.Bar_Cash = 1 Then
                    GoTo far
                End If
                With frmSales
                    Screen.MousePointer = 11
                    picHoldFocus.SetFocus
                    frmSales.Show
                    DoEvents
                    Me.Hide
                    Screen.MousePointer = 0
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
                        .grdMain.HighLight = flexHighlightAlways
                        .cmdInput(14).Caption = "Corr"
                    Else
                        .grdMain.HighLight = flexHighlightWithFocus
                        .cmdInput(14).Caption = "No Sale"
                    End If
                    .grdMain.Row = grdMain.Row
                    .grdMain.TopRow = grdMain.TopRow
                    If .grdMain.Rows = 1 Then
                        .cmdFancy(4).Caption = "Member No"
                    Else
                        .cmdFancy(4).Caption = "Discount"
                    End If
                    If grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Row, 0) = 12648447 Then
                        .cmdFancy(4).Caption = "Member No"
                    End If
                End With
            Else
far:
                With frmSales1
                    Screen.MousePointer = 11
                    picHoldFocus.SetFocus
                    Finalizing = False
                    frmSales1.Show
                    DoEvents
                    Me.Hide
                    Screen.MousePointer = 0
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
                        .grdMain.HighLight = flexHighlightAlways
                    Else
                        .grdMain.HighLight = flexHighlightWithFocus
                    End If
                    .grdMain.Row = grdMain.Row
                    .grdMain.TopRow = grdMain.TopRow
                End With
            End If
        Case 1
            Select Case UserRecord.uType
                Case 0 'Manager
                Case 1 'Night Manager
                Case 2 'Reservations Clerk
                Case 4 'Barman
                    If UserRecord.Reports = False Then
                        DisplayErr "Access Denied"
                        Exit Sub
                    End If
                Case 3 'Waiter
                Case 5 'GRV Clerck
                Case 6 'Buyer
                Case 7 'Supervisor
                Case 8 'Cashier
                    If UserRecord.Reports = False Then
                        DisplayErr "Access Denied"
                        Exit Sub
                    End If
                Case 9 'Owner
            End Select
            frmTillReport.Show
            DoEvents
            Me.Hide
    End Select
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
        On Error Resume Next
        picHoldFocus.SetFocus
        On Error GoTo 0
    End If
End Sub
Private Sub cmdSlip_Click()
    If grdMain.Enabled = False And cmdInput(Index).Caption <> "Void" Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    Select Case cmdSlip.Caption
        Case "View Slip"
            cmdSlip.Caption = "Close Slip"
            picSlip.Visible = True
            grdMain.SetFocus
        Case "Close Slip"
            cmdSlip.Caption = "View Slip"
            picSlip.Visible = False
    End Select
End Sub

Private Sub cmdSlip1_Click()
    Select Case cmdSlip1.Caption
        Case "Slip on"
            Select Case Panel_no
                Case 2
                    frmBar.cmdSlip1.Caption = "Slip off"
                    SaveSetting Trim(gblApp_Name), "Workstation", "BarSlip", 0
                Case 0
                    frmSales.cmdSlip.Caption = "Slip off"
                    SaveSetting Trim(gblApp_Name), "Workstation", "TillSlip", 0
            End Select
        Case "Slip off"
            Select Case Panel_no
                Case 2
                    frmBar.cmdSlip1.Caption = "Slip on"
                    SaveSetting Trim(gblApp_Name), "Workstation", "BarSlip", 1
                Case 0
                    frmSales.cmdSlip.Caption = "Slip on"
                    SaveSetting Trim(gblApp_Name), "Workstation", "TillSlip", 1
            End Select
    End Select
    send_data_steam_keylog (Me.Name & " - " & cmdSlip1.Caption)
End Sub
Private Sub errTimer_Timer()
    Select Case cmdErr.BackColor
        Case &HF2&      'White
            cmdErr.BackColor = &HFFFF&
        Case &HFFFF&    'Yellow
            cmdErr.BackColor = &HF2&
    End Select
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If cmdErr.Visible = True Then
        cmdErr.Visible = False
        cmdErr.BackColor = &HF2&
        errTimer.Enabled = False
        KeyRegister = ""
        lblKeyRegister = ""
        picDigit.Visible = False
        Exit Sub
    End If
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    Select Case KeyCode
        Case 38
            If picSlip.Visible = True Then
                grdMain.SetFocus
                DoEvents
            End If
        Case 40
            If picSlip.Visible = True Then
                grdMain.SetFocus
                DoEvents
            End If
        Case 13
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
        Case 120
            Key_Function cmdInput(12).Caption
        Case 121
            Key_Function cmdInput(13).Caption
        Case 122
            Key_Function cmdInput(16).Caption
        Case 123
            Key_Function cmdInput(21).Caption
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
                
                If InStr(KeyRegister, "Price O/V") <> 0 Then
                    For i = 1 To grdMain.Rows - 1
                        If Product_Code = grdMain.TextMatrix(i, 10) Then
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
            Key_Function "Dept"
    End Select
    KeyAscii = 0
End Sub

Public Sub LoadOldTab(Tab_No)
    TillData.ReturnTotal = 0
    TillData.UllageTotal = 0
    TillData.VoidTotal = 0
    TillData.Tendered = 0
    TillData.Cash = 0
    TillData.Card = 0
    TillData.Cheque = 0
    TillData.Charge = 0
    TillData.Loyalty = 0
    TillData.TaxTotal = 0
    TillData.TaxableSales = 0
    TillData.NonTaxableSales = 0
    TillData.CollectedTax = 0
    TillData.CalculatedTax = 0
    TillData.Corrects = 0
    TillData.DocNo = 0
    TillData.UserOveride = 0
    TillData.Discount = 0
    TillData.DiscountVal = 0
    TillData.Print_Count = 0
    ActiveReadServer "Select * from Tab_Listing where Tab_No= " & Tab_No & " order by Line_No"
    grdMain.Rows = 1
    grdMain.ColHidden(14) = True
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Trim(rs.Fields("Qty") & "")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Short_Desc")
        If Val(rs.Fields("Line_Total") & "") <> 0 Then
            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Line_Total"), "0.00")
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("KeyString")
        If Trim(rs.Fields("KeyString") & "") = "Subtotal" Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0FFC0
        End If
        If Trim(rs.Fields("KeyString") & "") = "" Then
            grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC00000
            grdMain.Cell(flexcpFontBold, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = True
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Cost")
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
        lblTab.Caption = "Tab: " & rs.Fields("Tab_Name")
        TillData.TabName = rs.Fields("Tab_Name")
        TillData.Covers = rs.Fields("Covers")
        TillData.DocNo = rs.Fields("Doc_No")
        TillData.Account_No = rs.Fields("Member_No") & ""
        rs.MoveNext
    Wend
    rs.Close
    ActiveReadServer ("Select count(*) as cnt from Print_Journal where Doc_no = " & TillData.DocNo & " and Doc_type = 'Bill Print' and Table_no = '" & TillData.TabNo & "'")
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
    lblKeyRegister.Caption = "Tab: " & TillData.TabName & " opened by " & UserRecord.Name
    If TillData.DocNo = 0 Then
        cmdKey(4).Caption = "No Sale"
    Else
        cmdKey(4).Caption = "Corr"
    End If
    ActiveUpdateServer "Update Tab_Listing set Locked =1 where Tab_No= " & Tab_No
    If frmInput1.cmdErr.Caption <> "Select a Tab to Print the Bill" Then
        GlobalMode = TillMode.Inputmode
        frmInput1.Hide
        If frmInput.cmdErr.Caption <> "Select a Tab to Close" Then
            frmBar.Show
        End If
    End If
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    chkdateagain = ""
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    If Me.Height < 10000 And newBack.Visible = False Then
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.782
            Me.Controls(i).Left = Me.Controls(i).Left * 0.782
            Me.Controls(i).Height = Me.Controls(i).Height * 0.79
            Me.Controls(i).top = Me.Controls(i).top * 0.79
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
        newBack.Width = Me.Width
        newBack.Height = Me.Height
    End If
    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.6
    grdMain.ColWidth(2) = grdMain.Width * 0.25
    grdMain.ColWidth(14) = 200
    Select Case GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="BarSlip", Default:=1)
        Case 0
            frmBar.cmdSlip1.Caption = "Slip off"
        Case 1
            frmBar.cmdSlip1.Caption = "Slip on"
    End Select
    Panel_no = 2
    If frmBar.Tag = "1" Then
        frmBar.Tag = ""
        Exit Sub
    End If
    dd = TillData.Account_No & "'"
    If picSlip.Visible = False Then cmdSlip.Caption = "View Slip"
    If picSlip.Visible = True Then cmdSlip.Caption = "Close Slip"
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    lblKeyRegister.TextAlign = fmTextAlignLeft
    lblKeyRegister.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    frmBar.KeyPreview = True
    grdDept.Rows = 0
    cmdDeptStrip(0).Caption = ""
    cmdDeptStrip(0).Picture = ""
    DoEvents
    Select Case Dept_Order
        Case 0
            ActiveReadServer "Select * from Departments_Panel3 where Location_no = '" & Location_No & "' ORDER BY Dept_Parent,Dept_Name"
        Case 1
            ActiveReadServer "Select * from Departments_Panel3 ORDER BY Dept_Parent,convert(int, substring(department_no,(SELECT PATINDEX('%-%', Department_No))+1,len(department_No)))"
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
    picHoldFocus.SetFocus
    If cmdPlu(0).Visible = False Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
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
       cmdPlu(b).Visible = False
    Next b
End Sub
Private Sub cmdPlu_Click(Index As Integer)
send_data_steam_keylog (Me.Name & " - " & cmdPlu(Index).Caption)
    DoEvents
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
            cmdPlu(b).Tag = ""
            cmdPlu(b).ToolTipText = ""
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
    On Error Resume Next
    picHoldFocus.SetFocus
    On Error GoTo 0
End Sub
Private Sub cmdDeptStrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdDeptStrip(Index).Picture <> App.Path & "\icons\downArr.bmp" And cmdDeptStrip(Index).Picture <> App.Path & "\icons\upArr.bmp" Then
        LoadPlu cmdDeptStrip(Index).Tag
    End If
    cmdSlip.Caption = "View Slip"
    picSlip.Visible = False
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
    On Error Resume Next
    picHoldFocus.SetFocus
    On Error GoTo 0
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If picSlip.Visible = True Then
        Select Case KeyCode
            Case 38
                If grdMain.Tag = "1" Then
                    grdMain.Tag = ""
                    If grdMain.Row > 1 Then grdMain.Row = grdMain.Row - 1
                End If
                If grdMain.Rows > 2 And grdMain.Row = 0 Then
                    grdMain.Row = 1
                End If
                grdMain.ShowCell grdMain.Row, 0
            Case 40
                If grdMain.Tag = "1" Then
                    grdMain.Tag = ""
                    If grdMain.Row <> grdMain.Rows - 1 Then grdMain.Row = grdMain.Row + 1
                End If
                If grdMain.Rows > 2 And grdMain.Row = 0 Then
                    grdMain.Row = 1
                End If
                grdMain.ShowCell grdMain.Row, 0
        End Select
    End If
    picHoldFocus.SetFocus
End Sub
Private Sub Form_Load()
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
    If Finalizing = True Then Exit Sub
    picSlip.Visible = False
    cmdSlip.Caption = "View Slip"
End Sub
Private Sub grdMain_GotFocus()
    grdMain.Tag = "1"
End Sub
Private Sub grdMain_LostFocus()
    grdMain.Tag = ""
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

Private Sub voidTimer_Timer()
    Select Case grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2)
        Case &HC0FFC0
            grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = &HFFC0C0
        Case &HFFC0C0
            grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = &HC0FFC0
    End Select
End Sub
