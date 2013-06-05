VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmValidate 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7530
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmValidate.frx":0000
   ScaleHeight     =   7530
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   11
      Left            =   4110
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   255
      Caption         =   "X"
      CaptionOffsetY  =   2
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   11
      Left            =   3390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5745
      Width           =   1515
      _Version        =   524298
      _ExtentX        =   2672
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "Ok"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   20.25
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":6148
      textLT          =   "frmValidate.frx":61AC
      textCT          =   "frmValidate.frx":61C4
      textRT          =   "frmValidate.frx":61DC
      textLM          =   "frmValidate.frx":61F4
      textRM          =   "frmValidate.frx":620C
      textLB          =   "frmValidate.frx":6224
      textCB          =   "frmValidate.frx":623C
      textRB          =   "frmValidate.frx":6254
      colorBack       =   "frmValidate.frx":626C
      colorIntern     =   "frmValidate.frx":6296
      colorMO         =   "frmValidate.frx":62C0
      colorFocus      =   "frmValidate.frx":62EA
      colorDisabled   =   "frmValidate.frx":6314
      colorPressed    =   "frmValidate.frx":633E
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   8
      Left            =   3390
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1515
      _Version        =   524298
      _ExtentX        =   2672
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "9"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":6368
      textLT          =   "frmValidate.frx":63CA
      textCT          =   "frmValidate.frx":63E2
      textRT          =   "frmValidate.frx":63FA
      textLM          =   "frmValidate.frx":6412
      textRM          =   "frmValidate.frx":642A
      textLB          =   "frmValidate.frx":6442
      textCB          =   "frmValidate.frx":645A
      textRB          =   "frmValidate.frx":6472
      colorBack       =   "frmValidate.frx":648A
      colorIntern     =   "frmValidate.frx":64B4
      colorMO         =   "frmValidate.frx":64DE
      colorFocus      =   "frmValidate.frx":6508
      colorDisabled   =   "frmValidate.frx":6532
      colorPressed    =   "frmValidate.frx":655C
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   2
      Left            =   3390
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1515
      _Version        =   524298
      _ExtentX        =   2672
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "3"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":6586
      textLT          =   "frmValidate.frx":65E8
      textCT          =   "frmValidate.frx":6600
      textRT          =   "frmValidate.frx":6618
      textLM          =   "frmValidate.frx":6630
      textRM          =   "frmValidate.frx":6648
      textLB          =   "frmValidate.frx":6660
      textCB          =   "frmValidate.frx":6678
      textRB          =   "frmValidate.frx":6690
      colorBack       =   "frmValidate.frx":66A8
      colorIntern     =   "frmValidate.frx":66D2
      colorMO         =   "frmValidate.frx":66FC
      colorFocus      =   "frmValidate.frx":6726
      colorDisabled   =   "frmValidate.frx":6750
      colorPressed    =   "frmValidate.frx":677A
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   5
      Left            =   3390
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2775
      Width           =   1515
      _Version        =   524298
      _ExtentX        =   2672
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "6"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":67A4
      textLT          =   "frmValidate.frx":6806
      textCT          =   "frmValidate.frx":681E
      textRT          =   "frmValidate.frx":6836
      textLM          =   "frmValidate.frx":684E
      textRM          =   "frmValidate.frx":6866
      textLB          =   "frmValidate.frx":687E
      textCB          =   "frmValidate.frx":6896
      textRB          =   "frmValidate.frx":68AE
      colorBack       =   "frmValidate.frx":68C6
      colorIntern     =   "frmValidate.frx":68F0
      colorMO         =   "frmValidate.frx":691A
      colorFocus      =   "frmValidate.frx":6944
      colorDisabled   =   "frmValidate.frx":696E
      colorPressed    =   "frmValidate.frx":6998
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   10
      Left            =   1845
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5745
      Width           =   1575
      _Version        =   524298
      _ExtentX        =   2778
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "0"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":69C2
      textLT          =   "frmValidate.frx":6A24
      textCT          =   "frmValidate.frx":6A3C
      textRT          =   "frmValidate.frx":6A54
      textLM          =   "frmValidate.frx":6A6C
      textRM          =   "frmValidate.frx":6A84
      textLB          =   "frmValidate.frx":6A9C
      textCB          =   "frmValidate.frx":6AB4
      textRB          =   "frmValidate.frx":6ACC
      colorBack       =   "frmValidate.frx":6AE4
      colorIntern     =   "frmValidate.frx":6B0E
      colorMO         =   "frmValidate.frx":6B38
      colorFocus      =   "frmValidate.frx":6B62
      colorDisabled   =   "frmValidate.frx":6B8C
      colorPressed    =   "frmValidate.frx":6BB6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   7
      Left            =   1845
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1575
      _Version        =   524298
      _ExtentX        =   2778
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "8"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":6BE0
      textLT          =   "frmValidate.frx":6C42
      textCT          =   "frmValidate.frx":6C5A
      textRT          =   "frmValidate.frx":6C72
      textLM          =   "frmValidate.frx":6C8A
      textRM          =   "frmValidate.frx":6CA2
      textLB          =   "frmValidate.frx":6CBA
      textCB          =   "frmValidate.frx":6CD2
      textRB          =   "frmValidate.frx":6CEA
      colorBack       =   "frmValidate.frx":6D02
      colorIntern     =   "frmValidate.frx":6D2C
      colorMO         =   "frmValidate.frx":6D56
      colorFocus      =   "frmValidate.frx":6D80
      colorDisabled   =   "frmValidate.frx":6DAA
      colorPressed    =   "frmValidate.frx":6DD4
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   1
      Left            =   1845
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "2"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":6DFE
      textLT          =   "frmValidate.frx":6E60
      textCT          =   "frmValidate.frx":6E78
      textRT          =   "frmValidate.frx":6E90
      textLM          =   "frmValidate.frx":6EA8
      textRM          =   "frmValidate.frx":6EC0
      textLB          =   "frmValidate.frx":6ED8
      textCB          =   "frmValidate.frx":6EF0
      textRB          =   "frmValidate.frx":6F08
      colorBack       =   "frmValidate.frx":6F20
      colorIntern     =   "frmValidate.frx":6F4A
      colorMO         =   "frmValidate.frx":6F74
      colorFocus      =   "frmValidate.frx":6F9E
      colorDisabled   =   "frmValidate.frx":6FC8
      colorPressed    =   "frmValidate.frx":6FF2
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   4
      Left            =   1845
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2775
      Width           =   1575
      _Version        =   524298
      _ExtentX        =   2778
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "5"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":701C
      textLT          =   "frmValidate.frx":707E
      textCT          =   "frmValidate.frx":7096
      textRT          =   "frmValidate.frx":70AE
      textLM          =   "frmValidate.frx":70C6
      textRM          =   "frmValidate.frx":70DE
      textLB          =   "frmValidate.frx":70F6
      textCB          =   "frmValidate.frx":710E
      textRB          =   "frmValidate.frx":7126
      colorBack       =   "frmValidate.frx":713E
      colorIntern     =   "frmValidate.frx":7168
      colorMO         =   "frmValidate.frx":7192
      colorFocus      =   "frmValidate.frx":71BC
      colorDisabled   =   "frmValidate.frx":71E6
      colorPressed    =   "frmValidate.frx":7210
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   9
      Left            =   300
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5745
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "CL"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":723A
      textLT          =   "frmValidate.frx":729E
      textCT          =   "frmValidate.frx":72B6
      textRT          =   "frmValidate.frx":72CE
      textLM          =   "frmValidate.frx":72E6
      textRM          =   "frmValidate.frx":72FE
      textLB          =   "frmValidate.frx":7316
      textCB          =   "frmValidate.frx":732E
      textRB          =   "frmValidate.frx":7346
      colorBack       =   "frmValidate.frx":735E
      colorIntern     =   "frmValidate.frx":7388
      colorMO         =   "frmValidate.frx":73B2
      colorFocus      =   "frmValidate.frx":73DC
      colorDisabled   =   "frmValidate.frx":7406
      colorPressed    =   "frmValidate.frx":7430
      Orientation     =   8
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   6
      Left            =   300
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "7"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":745A
      textLT          =   "frmValidate.frx":74BC
      textCT          =   "frmValidate.frx":74D4
      textRT          =   "frmValidate.frx":74EC
      textLM          =   "frmValidate.frx":7504
      textRM          =   "frmValidate.frx":751C
      textLB          =   "frmValidate.frx":7534
      textCB          =   "frmValidate.frx":754C
      textRB          =   "frmValidate.frx":7564
      colorBack       =   "frmValidate.frx":757C
      colorIntern     =   "frmValidate.frx":75A6
      colorMO         =   "frmValidate.frx":75D0
      colorFocus      =   "frmValidate.frx":75FA
      colorDisabled   =   "frmValidate.frx":7624
      colorPressed    =   "frmValidate.frx":764E
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   3
      Left            =   300
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2775
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "4"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":7678
      textLT          =   "frmValidate.frx":76DA
      textCT          =   "frmValidate.frx":76F2
      textRT          =   "frmValidate.frx":770A
      textLM          =   "frmValidate.frx":7722
      textRM          =   "frmValidate.frx":773A
      textLB          =   "frmValidate.frx":7752
      textCB          =   "frmValidate.frx":776A
      textRB          =   "frmValidate.frx":7782
      colorBack       =   "frmValidate.frx":779A
      colorIntern     =   "frmValidate.frx":77C4
      colorMO         =   "frmValidate.frx":77EE
      colorFocus      =   "frmValidate.frx":7818
      colorDisabled   =   "frmValidate.frx":7842
      colorPressed    =   "frmValidate.frx":786C
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   0
      Left            =   300
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2619
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
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
      Surface         =   1
      BackColorContainer=   1471145
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":7896
      textLT          =   "frmValidate.frx":78F8
      textCT          =   "frmValidate.frx":7910
      textRT          =   "frmValidate.frx":7928
      textLM          =   "frmValidate.frx":7940
      textRM          =   "frmValidate.frx":7958
      textLB          =   "frmValidate.frx":7970
      textCB          =   "frmValidate.frx":7988
      textRB          =   "frmValidate.frx":79A0
      colorBack       =   "frmValidate.frx":79B8
      colorIntern     =   "frmValidate.frx":79E2
      colorMO         =   "frmValidate.frx":7A0C
      colorFocus      =   "frmValidate.frx":7A36
      colorDisabled   =   "frmValidate.frx":7A60
      colorPressed    =   "frmValidate.frx":7A8A
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   915
      Index           =   1
      Left            =   5520
      TabIndex        =   15
      Top             =   1785
      Visible         =   0   'False
      Width           =   3960
      _Version        =   524298
      _ExtentX        =   6985
      _ExtentY        =   1614
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":7AB4
      textLT          =   "frmValidate.frx":7B16
      textCT          =   "frmValidate.frx":7B2E
      textRT          =   "frmValidate.frx":7B46
      textLM          =   "frmValidate.frx":7B5E
      textRM          =   "frmValidate.frx":7B76
      textLB          =   "frmValidate.frx":7B8E
      textCB          =   "frmValidate.frx":7BA6
      textRB          =   "frmValidate.frx":7BBE
      colorBack       =   "frmValidate.frx":7BD6
      colorIntern     =   "frmValidate.frx":7C00
      colorMO         =   "frmValidate.frx":7C2A
      colorFocus      =   "frmValidate.frx":7C54
      colorDisabled   =   "frmValidate.frx":7C7E
      colorPressed    =   "frmValidate.frx":7CA8
      Style           =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   915
      Index           =   2
      Left            =   5520
      TabIndex        =   16
      Top             =   2700
      Visible         =   0   'False
      Width           =   3960
      _Version        =   524298
      _ExtentX        =   6985
      _ExtentY        =   1614
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":7CD2
      textLT          =   "frmValidate.frx":7D34
      textCT          =   "frmValidate.frx":7D4C
      textRT          =   "frmValidate.frx":7D64
      textLM          =   "frmValidate.frx":7D7C
      textRM          =   "frmValidate.frx":7D94
      textLB          =   "frmValidate.frx":7DAC
      textCB          =   "frmValidate.frx":7DC4
      textRB          =   "frmValidate.frx":7DDC
      colorBack       =   "frmValidate.frx":7DF4
      colorIntern     =   "frmValidate.frx":7E1E
      colorMO         =   "frmValidate.frx":7E48
      colorFocus      =   "frmValidate.frx":7E72
      colorDisabled   =   "frmValidate.frx":7E9C
      colorPressed    =   "frmValidate.frx":7EC6
      Style           =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   915
      Index           =   3
      Left            =   5520
      TabIndex        =   17
      Top             =   3615
      Visible         =   0   'False
      Width           =   3960
      _Version        =   524298
      _ExtentX        =   6985
      _ExtentY        =   1614
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":7EF0
      textLT          =   "frmValidate.frx":7F52
      textCT          =   "frmValidate.frx":7F6A
      textRT          =   "frmValidate.frx":7F82
      textLM          =   "frmValidate.frx":7F9A
      textRM          =   "frmValidate.frx":7FB2
      textLB          =   "frmValidate.frx":7FCA
      textCB          =   "frmValidate.frx":7FE2
      textRB          =   "frmValidate.frx":7FFA
      colorBack       =   "frmValidate.frx":8012
      colorIntern     =   "frmValidate.frx":803C
      colorMO         =   "frmValidate.frx":8066
      colorFocus      =   "frmValidate.frx":8090
      colorDisabled   =   "frmValidate.frx":80BA
      colorPressed    =   "frmValidate.frx":80E4
      Style           =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   915
      Index           =   4
      Left            =   5520
      TabIndex        =   18
      Top             =   4530
      Visible         =   0   'False
      Width           =   3960
      _Version        =   524298
      _ExtentX        =   6985
      _ExtentY        =   1614
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      SpecialEffect   =   1
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":810E
      textLT          =   "frmValidate.frx":8170
      textCT          =   "frmValidate.frx":8188
      textRT          =   "frmValidate.frx":81A0
      textLM          =   "frmValidate.frx":81B8
      textRM          =   "frmValidate.frx":81D0
      textLB          =   "frmValidate.frx":81E8
      textCB          =   "frmValidate.frx":8200
      textRB          =   "frmValidate.frx":8218
      colorBack       =   "frmValidate.frx":8230
      colorIntern     =   "frmValidate.frx":825A
      colorMO         =   "frmValidate.frx":8284
      colorFocus      =   "frmValidate.frx":82AE
      colorDisabled   =   "frmValidate.frx":82D8
      colorPressed    =   "frmValidate.frx":8302
      Style           =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   915
      Index           =   5
      Left            =   5520
      TabIndex        =   19
      Top             =   5445
      Visible         =   0   'False
      Width           =   3960
      _Version        =   524298
      _ExtentX        =   6985
      _ExtentY        =   1614
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":832C
      textLT          =   "frmValidate.frx":838E
      textCT          =   "frmValidate.frx":83A6
      textRT          =   "frmValidate.frx":83BE
      textLM          =   "frmValidate.frx":83D6
      textRM          =   "frmValidate.frx":83EE
      textLB          =   "frmValidate.frx":8406
      textCB          =   "frmValidate.frx":841E
      textRB          =   "frmValidate.frx":8436
      colorBack       =   "frmValidate.frx":844E
      colorIntern     =   "frmValidate.frx":8478
      colorMO         =   "frmValidate.frx":84A2
      colorFocus      =   "frmValidate.frx":84CC
      colorDisabled   =   "frmValidate.frx":84F6
      colorPressed    =   "frmValidate.frx":8520
      Style           =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   885
      Index           =   6
      Left            =   5520
      TabIndex        =   20
      Top             =   6360
      Visible         =   0   'False
      Width           =   3960
      _Version        =   524298
      _ExtentX        =   6985
      _ExtentY        =   1561
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      CornerFactor    =   15
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":854A
      textLT          =   "frmValidate.frx":85AC
      textCT          =   "frmValidate.frx":85C4
      textRT          =   "frmValidate.frx":85DC
      textLM          =   "frmValidate.frx":85F4
      textRM          =   "frmValidate.frx":860C
      textLB          =   "frmValidate.frx":8624
      textCB          =   "frmValidate.frx":863C
      textRB          =   "frmValidate.frx":8654
      colorBack       =   "frmValidate.frx":866C
      colorIntern     =   "frmValidate.frx":8696
      colorMO         =   "frmValidate.frx":86C0
      colorFocus      =   "frmValidate.frx":86EA
      colorDisabled   =   "frmValidate.frx":8714
      colorPressed    =   "frmValidate.frx":873E
      Style           =   2
      Orientation     =   4
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   915
      Index           =   0
      Left            =   5520
      TabIndex        =   21
      Top             =   870
      Visible         =   0   'False
      Width           =   3960
      _Version        =   524298
      _ExtentX        =   6985
      _ExtentY        =   1614
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":8768
      textLT          =   "frmValidate.frx":87CA
      textCT          =   "frmValidate.frx":87E2
      textRT          =   "frmValidate.frx":87FA
      textLM          =   "frmValidate.frx":8812
      textRM          =   "frmValidate.frx":882A
      textLB          =   "frmValidate.frx":8842
      textCB          =   "frmValidate.frx":885A
      textRB          =   "frmValidate.frx":8872
      colorBack       =   "frmValidate.frx":888A
      colorIntern     =   "frmValidate.frx":88B4
      colorMO         =   "frmValidate.frx":88DE
      colorFocus      =   "frmValidate.frx":8908
      colorDisabled   =   "frmValidate.frx":8932
      colorPressed    =   "frmValidate.frx":895C
      Style           =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTop 
      Height          =   615
      Left            =   5520
      TabIndex        =   22
      Top             =   270
      Width           =   3960
      _Version        =   524298
      _ExtentX        =   6985
      _ExtentY        =   1085
      _StockProps     =   66
      Caption         =   "Select an Override Reason"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      CornerFactor    =   15
      BackColorContainer=   4498384
      SpecialEffect   =   1
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmValidate.frx":8986
      textLT          =   "frmValidate.frx":8A18
      textCT          =   "frmValidate.frx":8A30
      textRT          =   "frmValidate.frx":8A48
      textLM          =   "frmValidate.frx":8A60
      textRM          =   "frmValidate.frx":8A78
      textLB          =   "frmValidate.frx":8A90
      textCB          =   "frmValidate.frx":8AA8
      textRB          =   "frmValidate.frx":8AC0
      colorBack       =   "frmValidate.frx":8AD8
      colorIntern     =   "frmValidate.frx":8B02
      colorMO         =   "frmValidate.frx":8B2C
      colorFocus      =   "frmValidate.frx":8B56
      colorDisabled   =   "frmValidate.frx":8B80
      colorPressed    =   "frmValidate.frx":8BAA
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   630
      TabIndex        =   14
      Top             =   540
      Width           =   2295
   End
   Begin MSForms.TextBox txtPass 
      Height          =   585
      Left            =   480
      TabIndex        =   0
      Top             =   390
      Visible         =   0   'False
      Width           =   2685
      VariousPropertyBits=   746604563
      ForeColor       =   16777215
      Size            =   "4736;1032"
      PasswordChar    =   42
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   555
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdInput_Click(Index As Integer)
    send_data_steam_keylog (Me.Name & " - " & frmValidate.Tag & " - " & cmdInput(Index).Caption)
    Reason_No = 0
    Select Case cmdInput(Index).Caption
        Case "Ok"
            If frmValidate.Width = 9945 Then
                ReasonSelect = False
                For i = 0 To cmdCook.Count - 1
                    If cmdCook(i).Value = 1 Then
                        Reason_No = i + 1
                        ReasonSelect = True
                        Exit For
                    End If
                Next i
                Select Case ReasonSelect
                    Case True
                    Case False
                        Load frmError
                        frmError.Caption = " Override Reason"
                        frmError.lblCap.Caption = "You have not Selected a Reason for this Override"
                        frmError.lblError.Caption = err.Description
                        DoEvents
                        frmError.Show vbModal
                        Exit Sub
                End Select
            End If
            Select Case frmValidate.Tag
                Case "Discount"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Disc_Perc = 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = Val(rs.Fields("User_No") & "." & Reason_No)
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                Case "Corr"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Item_Corrects = 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = Val(rs.Fields("User_No") & "." & Reason_No)
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                Case "Void"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Voids = 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = Val(rs.Fields("User_No") & "." & Reason_No)
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                Case "Return"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Returns = 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = Val(rs.Fields("User_No") & "." & Reason_No)
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                Case "Wastage"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Ullages = 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = Val(rs.Fields("User_No") & "." & Reason_No)
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                Case "Pay Out"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Payouts = 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = Val(rs.Fields("User_No") & "." & Reason_No)
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                 Case "Price O/V"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Override= 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = Val(rs.Fields("User_No") & "." & Reason_No)
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                 Case "Service Charge"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Service_Charge = 1"
                    If rs.RecordCount > 0 Then

                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                 
                  Case "Quotes"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Quotes = 1"
                    If rs.RecordCount > 0 Then

                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                 Case "Reprint", "Print Bill"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Reprint = 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = rs.Fields("User_No")
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                    rs.Close
                 Case "Owner Transfer"
                    ActiveReadServer "Select User_No from Users where User_Password= '" & txtPass.Text & "' and Owner_transfer = 1"
                    If rs.RecordCount > 0 Then
                        TillData.UserOveride = rs.Fields("User_No")
                        frmValidate.Tag = "1"
                    Else
                        frmValidate.Tag = "0"
                    End If
                Case ""
            End Select
            Me.Hide
        Case "CL"
            txtPass.Visible = False
            lblPassword.Visible = True
            txtPass.Text = ""
        Case Else
            lblPassword.Visible = False
            txtPass.Visible = True
            txtPass.SetFocus
            SendKeys cmdInput(Index).Caption
    End Select
    send_data_steam_keylog (Me.Name & " - Validated by user " & TillData.UserOveride)
End Sub
Private Sub cmdKey_Click(Index As Integer)
    frmValidate.Tag = ""
    Me.Hide
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    If frmValidate.Tag = "Service Charge" Then
        frmValidate.Width = 5265
    End If
    txtPass.Text = ""
    Select Case Panel_no
        Case 0: frmSales.Tag = "1"
        Case 1: frmSales1.Tag = "1"
        Case 2: frmBar.Tag = "1"
    End Select
    lblPassword.Visible = False
    txtPass.Visible = True
    txtPass.SetFocus
    For i = 0 To cmdCook.Count - 1
        cmdCook(i).Visible = False
        cmdCook(i).Value = 0
    Next i
    i = -1
    ActiveReadServer "Select * from Reasons order by Reason_No"
    While Not rs.EOF
        i = i + 1
        cmdCook(i).Caption = rs.Fields("Reason_Name")
        cmdCook(i).Visible = True
        rs.MoveNext
    Wend
    rs.Close
    On Error GoTo 0
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
            txtPass.Text = ""
            lblPassword.Visible = True
        Case 46, 48 To 57
            lblPassword.Visible = False
            txtPass.Visible = True
            txtPass.SetFocus
            If txtPass.Text <> "" Then
                If Asc(Mid(txtPass.Text, 1, 1)) < 48 Or Asc(Mid(txtPass.Text, 1, 1)) > 57 Then
                    If txtPass.Text = "." Then
                       txtPass.Text = txtPass.Text & Chr(KeyAscii)
                    Else
                        If InStr(txtPass.Text, ".") = 0 Then
                            txtPass.Text = Chr(KeyAscii)
                        Else
                            If Left(txtPass.Text, 7) = "Counted" Then
                                txtPass.Text = Chr(KeyAscii)
                            Else
                                txtPass.Text = txtPass.Text & Chr(KeyAscii)
                            End If
                        End If
                    End If
                Else
                    txtPass.Text = txtPass.Text & Chr(KeyAscii)
                End If
            Else
                txtPass.Text = txtPass.Text & Chr(KeyAscii)
            End If
            KeyAscii = 0
    End Select
End Sub
Private Sub Form_Load()
    Select Case VoidReasons
        Case 0: frmValidate.Width = 5265
        Case 1: frmValidate.Width = 9855
    End Select
End Sub
Private Sub txtPass_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
       KeyCode = 0
       cmdInput_Click 11
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    Case 8, 27, 48 To 57
    Case Else
        KeyAscii = 0
End Select

End Sub
