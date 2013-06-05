VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmDiscount 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7530
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDiscount.frx":0000
   ScaleHeight     =   7530
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer errTimer 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1515
      Index           =   9
      Left            =   2130
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2672
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":5B7A
      textLT          =   "frmDiscount.frx":5BDE
      textCT          =   "frmDiscount.frx":5BF6
      textRT          =   "frmDiscount.frx":5C0E
      textLM          =   "frmDiscount.frx":5C26
      textRM          =   "frmDiscount.frx":5C3E
      textLB          =   "frmDiscount.frx":5C56
      textCB          =   "frmDiscount.frx":5C6E
      textRB          =   "frmDiscount.frx":5C86
      colorBack       =   "frmDiscount.frx":5C9E
      colorIntern     =   "frmDiscount.frx":5CC8
      colorMO         =   "frmDiscount.frx":5CF2
      colorFocus      =   "frmDiscount.frx":5D1C
      colorDisabled   =   "frmDiscount.frx":5D46
      colorPressed    =   "frmDiscount.frx":5D70
      Orientation     =   8
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   0
      Left            =   2130
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1260
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":5D9A
      textLT          =   "frmDiscount.frx":5DFC
      textCT          =   "frmDiscount.frx":5E14
      textRT          =   "frmDiscount.frx":5E2C
      textLM          =   "frmDiscount.frx":5E44
      textRM          =   "frmDiscount.frx":5E5C
      textLB          =   "frmDiscount.frx":5E74
      textCB          =   "frmDiscount.frx":5E8C
      textRB          =   "frmDiscount.frx":5EA4
      colorBack       =   "frmDiscount.frx":5EBC
      colorIntern     =   "frmDiscount.frx":5EE6
      colorMO         =   "frmDiscount.frx":5F10
      colorFocus      =   "frmDiscount.frx":5F3A
      colorDisabled   =   "frmDiscount.frx":5F64
      colorPressed    =   "frmDiscount.frx":5F8E
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1515
      Index           =   10
      Left            =   3645
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   2672
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
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":5FB8
      textLT          =   "frmDiscount.frx":601A
      textCT          =   "frmDiscount.frx":6032
      textRT          =   "frmDiscount.frx":604A
      textLM          =   "frmDiscount.frx":6062
      textRM          =   "frmDiscount.frx":607A
      textLB          =   "frmDiscount.frx":6092
      textCB          =   "frmDiscount.frx":60AA
      textRB          =   "frmDiscount.frx":60C2
      colorBack       =   "frmDiscount.frx":60DA
      colorIntern     =   "frmDiscount.frx":6104
      colorMO         =   "frmDiscount.frx":612E
      colorFocus      =   "frmDiscount.frx":6158
      colorDisabled   =   "frmDiscount.frx":6182
      colorPressed    =   "frmDiscount.frx":61AC
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   1
      Left            =   3645
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
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
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":61D6
      textLT          =   "frmDiscount.frx":6238
      textCT          =   "frmDiscount.frx":6250
      textRT          =   "frmDiscount.frx":6268
      textLM          =   "frmDiscount.frx":6280
      textRM          =   "frmDiscount.frx":6298
      textLB          =   "frmDiscount.frx":62B0
      textCB          =   "frmDiscount.frx":62C8
      textRB          =   "frmDiscount.frx":62E0
      colorBack       =   "frmDiscount.frx":62F8
      colorIntern     =   "frmDiscount.frx":6322
      colorMO         =   "frmDiscount.frx":634C
      colorFocus      =   "frmDiscount.frx":6376
      colorDisabled   =   "frmDiscount.frx":63A0
      colorPressed    =   "frmDiscount.frx":63CA
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   11
      Left            =   8070
      TabIndex        =   4
      Top             =   270
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   2163158
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
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1500
      Index           =   0
      Left            =   300
      TabIndex        =   5
      Top             =   1260
      Width           =   1695
      _Version        =   524298
      _ExtentX        =   2990
      _ExtentY        =   2646
      _StockProps     =   66
      Caption         =   "Amount"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":63F4
      textLT          =   "frmDiscount.frx":6460
      textCT          =   "frmDiscount.frx":6478
      textRT          =   "frmDiscount.frx":6490
      textLM          =   "frmDiscount.frx":64A8
      textRM          =   "frmDiscount.frx":64C0
      textLB          =   "frmDiscount.frx":64D8
      textCB          =   "frmDiscount.frx":64F0
      textRB          =   "frmDiscount.frx":6508
      colorBack       =   "frmDiscount.frx":6520
      colorIntern     =   "frmDiscount.frx":654A
      colorMO         =   "frmDiscount.frx":6574
      colorFocus      =   "frmDiscount.frx":659E
      colorDisabled   =   "frmDiscount.frx":65C8
      colorPressed    =   "frmDiscount.frx":65F2
      Style           =   2
      Orientation     =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1470
      Index           =   1
      Left            =   300
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
      _Version        =   524298
      _ExtentX        =   2990
      _ExtentY        =   2593
      _StockProps     =   66
      Caption         =   "%"
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":661C
      textLT          =   "frmDiscount.frx":667E
      textCT          =   "frmDiscount.frx":6696
      textRT          =   "frmDiscount.frx":66AE
      textLM          =   "frmDiscount.frx":66C6
      textRM          =   "frmDiscount.frx":66DE
      textLB          =   "frmDiscount.frx":66F6
      textCB          =   "frmDiscount.frx":670E
      textRB          =   "frmDiscount.frx":6726
      colorBack       =   "frmDiscount.frx":673E
      colorIntern     =   "frmDiscount.frx":6768
      colorMO         =   "frmDiscount.frx":6792
      colorFocus      =   "frmDiscount.frx":67BC
      colorDisabled   =   "frmDiscount.frx":67E6
      colorPressed    =   "frmDiscount.frx":6810
      Style           =   2
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1530
      Index           =   11
      Left            =   300
      TabIndex        =   7
      Top             =   5700
      Width           =   1695
      _Version        =   524298
      _ExtentX        =   2990
      _ExtentY        =   2699
      _StockProps     =   66
      Caption         =   "Ok"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   18
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":683A
      textLT          =   "frmDiscount.frx":689E
      textCT          =   "frmDiscount.frx":68B6
      textRT          =   "frmDiscount.frx":68CE
      textLM          =   "frmDiscount.frx":68E6
      textRM          =   "frmDiscount.frx":68FE
      textLB          =   "frmDiscount.frx":6916
      textCB          =   "frmDiscount.frx":692E
      textRB          =   "frmDiscount.frx":6946
      colorBack       =   "frmDiscount.frx":695E
      colorIntern     =   "frmDiscount.frx":6988
      colorMO         =   "frmDiscount.frx":69B2
      colorFocus      =   "frmDiscount.frx":69DC
      colorDisabled   =   "frmDiscount.frx":6A06
      colorPressed    =   "frmDiscount.frx":6A30
      Orientation     =   4
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   2
      Left            =   5220
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1575
      _Version        =   524298
      _ExtentX        =   2778
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":6A5A
      textLT          =   "frmDiscount.frx":6ABC
      textCT          =   "frmDiscount.frx":6AD4
      textRT          =   "frmDiscount.frx":6AEC
      textLM          =   "frmDiscount.frx":6B04
      textRM          =   "frmDiscount.frx":6B1C
      textLB          =   "frmDiscount.frx":6B34
      textCB          =   "frmDiscount.frx":6B4C
      textRB          =   "frmDiscount.frx":6B64
      colorBack       =   "frmDiscount.frx":6B7C
      colorIntern     =   "frmDiscount.frx":6BA6
      colorMO         =   "frmDiscount.frx":6BD0
      colorFocus      =   "frmDiscount.frx":6BFA
      colorDisabled   =   "frmDiscount.frx":6C24
      colorPressed    =   "frmDiscount.frx":6C4E
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdDot 
      Height          =   1515
      Left            =   5220
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5715
      Width           =   1575
      _Version        =   524298
      _ExtentX        =   2778
      _ExtentY        =   2672
      _StockProps     =   66
      Caption         =   "."
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   36
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
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":6C78
      textLT          =   "frmDiscount.frx":6CDA
      textCT          =   "frmDiscount.frx":6CF2
      textRT          =   "frmDiscount.frx":6D0A
      textLM          =   "frmDiscount.frx":6D22
      textRM          =   "frmDiscount.frx":6D3A
      textLB          =   "frmDiscount.frx":6D52
      textCB          =   "frmDiscount.frx":6D6A
      textRB          =   "frmDiscount.frx":6D82
      colorBack       =   "frmDiscount.frx":6D9A
      colorIntern     =   "frmDiscount.frx":6DC4
      colorMO         =   "frmDiscount.frx":6DEE
      colorFocus      =   "frmDiscount.frx":6E18
      colorDisabled   =   "frmDiscount.frx":6E42
      colorPressed    =   "frmDiscount.frx":6E6C
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   885
      Left            =   240
      TabIndex        =   10
      Top             =   270
      Visible         =   0   'False
      Width           =   7665
      _Version        =   524298
      _ExtentX        =   13520
      _ExtentY        =   1561
      _StockProps     =   66
      Caption         =   "Invalid Key Pressed"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":6E96
      textLT          =   "frmDiscount.frx":6F1C
      textCT          =   "frmDiscount.frx":6F34
      textRT          =   "frmDiscount.frx":6F4C
      textLM          =   "frmDiscount.frx":6F64
      textRM          =   "frmDiscount.frx":6F7C
      textLB          =   "frmDiscount.frx":6F94
      textCB          =   "frmDiscount.frx":6FAC
      textRB          =   "frmDiscount.frx":6FC4
      colorBack       =   "frmDiscount.frx":6FDC
      colorIntern     =   "frmDiscount.frx":7006
      colorMO         =   "frmDiscount.frx":7030
      colorFocus      =   "frmDiscount.frx":705A
      colorDisabled   =   "frmDiscount.frx":7084
      colorPressed    =   "frmDiscount.frx":70AE
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   0
      Left            =   6960
      TabIndex        =   13
      Top             =   1290
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "-10%"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   1
      Left            =   6960
      TabIndex        =   14
      Top             =   2475
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "-20%"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   2
      Left            =   6960
      TabIndex        =   15
      Top             =   3660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "-30%"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   3
      Left            =   6960
      TabIndex        =   16
      Top             =   4845
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "-40%"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   4
      Left            =   6960
      TabIndex        =   17
      Top             =   6030
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "-50%"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   5
      Left            =   7920
      TabIndex        =   18
      Top             =   1290
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "-60%"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   6
      Left            =   7920
      TabIndex        =   19
      Top             =   2475
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "-70%"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   7
      Left            =   7920
      TabIndex        =   20
      Top             =   3660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "-80%"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   8
      Left            =   7920
      TabIndex        =   21
      Top             =   4845
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   11329524
      Caption         =   "Cost +10"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1215
      Index           =   9
      Left            =   7920
      TabIndex        =   22
      Top             =   6030
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      Appearance      =   3
      BackColor       =   1610181
      Caption         =   "Free"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BTNENHLib4.BtnEnh cmdApply 
      Height          =   1470
      Left            =   300
      TabIndex        =   23
      Top             =   4230
      Width           =   1695
      _Version        =   524298
      _ExtentX        =   2990
      _ExtentY        =   2593
      _StockProps     =   66
      Caption         =   "Whole Sale"
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":70D8
      textLT          =   "frmDiscount.frx":714C
      textCT          =   "frmDiscount.frx":7164
      textRT          =   "frmDiscount.frx":717C
      textLM          =   "frmDiscount.frx":7194
      textRM          =   "frmDiscount.frx":71AC
      textLB          =   "frmDiscount.frx":71C4
      textCB          =   "frmDiscount.frx":71DC
      textRB          =   "frmDiscount.frx":71F4
      colorBack       =   "frmDiscount.frx":720C
      colorIntern     =   "frmDiscount.frx":7236
      colorMO         =   "frmDiscount.frx":7260
      colorFocus      =   "frmDiscount.frx":728A
      colorDisabled   =   "frmDiscount.frx":72B4
      colorPressed    =   "frmDiscount.frx":72DE
      Orientation     =   4
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   6
      Left            =   2130
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4230
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
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":7308
      textLT          =   "frmDiscount.frx":736A
      textCT          =   "frmDiscount.frx":7382
      textRT          =   "frmDiscount.frx":739A
      textLM          =   "frmDiscount.frx":73B2
      textRM          =   "frmDiscount.frx":73CA
      textLB          =   "frmDiscount.frx":73E2
      textCB          =   "frmDiscount.frx":73FA
      textRB          =   "frmDiscount.frx":7412
      colorBack       =   "frmDiscount.frx":742A
      colorIntern     =   "frmDiscount.frx":7454
      colorMO         =   "frmDiscount.frx":747E
      colorFocus      =   "frmDiscount.frx":74A8
      colorDisabled   =   "frmDiscount.frx":74D2
      colorPressed    =   "frmDiscount.frx":74FC
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   3
      Left            =   2130
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2745
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
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":7526
      textLT          =   "frmDiscount.frx":7588
      textCT          =   "frmDiscount.frx":75A0
      textRT          =   "frmDiscount.frx":75B8
      textLM          =   "frmDiscount.frx":75D0
      textRM          =   "frmDiscount.frx":75E8
      textLB          =   "frmDiscount.frx":7600
      textCB          =   "frmDiscount.frx":7618
      textRB          =   "frmDiscount.frx":7630
      colorBack       =   "frmDiscount.frx":7648
      colorIntern     =   "frmDiscount.frx":7672
      colorMO         =   "frmDiscount.frx":769C
      colorFocus      =   "frmDiscount.frx":76C6
      colorDisabled   =   "frmDiscount.frx":76F0
      colorPressed    =   "frmDiscount.frx":771A
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   7
      Left            =   3645
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4230
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
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
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":7744
      textLT          =   "frmDiscount.frx":77A6
      textCT          =   "frmDiscount.frx":77BE
      textRT          =   "frmDiscount.frx":77D6
      textLM          =   "frmDiscount.frx":77EE
      textRM          =   "frmDiscount.frx":7806
      textLB          =   "frmDiscount.frx":781E
      textCB          =   "frmDiscount.frx":7836
      textRB          =   "frmDiscount.frx":784E
      colorBack       =   "frmDiscount.frx":7866
      colorIntern     =   "frmDiscount.frx":7890
      colorMO         =   "frmDiscount.frx":78BA
      colorFocus      =   "frmDiscount.frx":78E4
      colorDisabled   =   "frmDiscount.frx":790E
      colorPressed    =   "frmDiscount.frx":7938
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   4
      Left            =   3645
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2745
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
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
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":7962
      textLT          =   "frmDiscount.frx":79C4
      textCT          =   "frmDiscount.frx":79DC
      textRT          =   "frmDiscount.frx":79F4
      textLM          =   "frmDiscount.frx":7A0C
      textRM          =   "frmDiscount.frx":7A24
      textLB          =   "frmDiscount.frx":7A3C
      textCB          =   "frmDiscount.frx":7A54
      textRB          =   "frmDiscount.frx":7A6C
      colorBack       =   "frmDiscount.frx":7A84
      colorIntern     =   "frmDiscount.frx":7AAE
      colorMO         =   "frmDiscount.frx":7AD8
      colorFocus      =   "frmDiscount.frx":7B02
      colorDisabled   =   "frmDiscount.frx":7B2C
      colorPressed    =   "frmDiscount.frx":7B56
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   5
      Left            =   5220
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2745
      Width           =   1575
      _Version        =   524298
      _ExtentX        =   2778
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
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":7B80
      textLT          =   "frmDiscount.frx":7BE2
      textCT          =   "frmDiscount.frx":7BFA
      textRT          =   "frmDiscount.frx":7C12
      textLM          =   "frmDiscount.frx":7C2A
      textRM          =   "frmDiscount.frx":7C42
      textLB          =   "frmDiscount.frx":7C5A
      textCB          =   "frmDiscount.frx":7C72
      textRB          =   "frmDiscount.frx":7C8A
      colorBack       =   "frmDiscount.frx":7CA2
      colorIntern     =   "frmDiscount.frx":7CCC
      colorMO         =   "frmDiscount.frx":7CF6
      colorFocus      =   "frmDiscount.frx":7D20
      colorDisabled   =   "frmDiscount.frx":7D4A
      colorPressed    =   "frmDiscount.frx":7D74
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1485
      Index           =   8
      Left            =   5220
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4230
      Width           =   1575
      _Version        =   524298
      _ExtentX        =   2778
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
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmDiscount.frx":7D9E
      textLT          =   "frmDiscount.frx":7E00
      textCT          =   "frmDiscount.frx":7E18
      textRT          =   "frmDiscount.frx":7E30
      textLM          =   "frmDiscount.frx":7E48
      textRM          =   "frmDiscount.frx":7E60
      textLB          =   "frmDiscount.frx":7E78
      textCB          =   "frmDiscount.frx":7E90
      textRB          =   "frmDiscount.frx":7EA8
      colorBack       =   "frmDiscount.frx":7EC0
      colorIntern     =   "frmDiscount.frx":7EEA
      colorMO         =   "frmDiscount.frx":7F14
      colorFocus      =   "frmDiscount.frx":7F3E
      colorDisabled   =   "frmDiscount.frx":7F68
      colorPressed    =   "frmDiscount.frx":7F92
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin MSForms.Label lblHeading 
      Height          =   525
      Left            =   2010
      TabIndex        =   12
      Top             =   540
      Width           =   2295
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "4048;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   405
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Left            =   510
      TabIndex        =   11
      Top             =   540
      Width           =   1485
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Discount:"
      Size            =   "2619;873"
      FontName        =   "Arial Narrow"
      FontHeight      =   405
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdApply_Click()
    Select Case cmdApply.Caption
        Case "Whole Sale"
            cmdApply.Caption = "Current Line Only"
        Case "Current Line Only"
            cmdApply.Caption = "Whole Sale"
    End Select
End Sub

Private Sub cmdDeptStrip_Click(Index As Integer)
    If Val(lblHeading.Caption) > 99.99 And Index = 1 Then
        cmdErr.Caption = "Percentage discount cannot exceed 100%."
        cmdErr.Visible = True
        errTimer.Enabled = True
        lblHeading.Caption = "0.00"
        Exit Sub
    End If
    If Val(lblHeading.Caption) > TillData.SaleTotal And Index = 0 Then
        cmdErr.Caption = "Discount value cannot exceed the sale value."
        cmdErr.Visible = True
        errTimer.Enabled = True
        lblHeading.Caption = "0.00"
        Exit Sub
    End If
End Sub

Private Sub cmdDot_Click()
    If InStr(lblHeading.Caption, ".") <> 0 Then Exit Sub
    lblHeading.Caption = lblHeading.Caption & "."
End Sub

Private Sub cmdErr_Click()
    lblHeading = "0.00"
    cmdErr.Visible = False
    errTimer.Enabled = False
End Sub
Private Sub cmdInput_Click(Index As Integer)
    Select Case cmdInput(Index).Caption
        Case "0" To "9"
            If cmdErr.Visible = True Then
                lblHeading = "0.00"
                cmdErr.Visible = False
                errTimer.Enabled = False
            End If
            If lblHeading.Caption = "0.00" Then lblHeading = ""
            lblHeading.Caption = lblHeading.Caption & cmdInput(Index).Caption
        Case "Ok"
            If cmdErr.Visible = True Then Exit Sub
            If Val(lblHeading.Caption) = 0 Then
                cmdErr.Caption = "Amount or Percentage Required"
                cmdErr.Visible = True
                errTimer.Enabled = True
                lblHeading.Caption = "0.00"
                Exit Sub
            End If
            If cmdDeptStrip(0).Value = 0 And cmdDeptStrip(1).Value = 0 Then
                cmdErr.Caption = "Amount or Percentage Required"
                cmdErr.Visible = True
                errTimer.Enabled = True
                lblHeading.Caption = "0.00"
                Exit Sub
            End If
            If cmdDeptStrip(0).Value = 1 Then
                
                TillData.DiscountVal = Val(lblHeading.Caption)
            Else
                
                TillData.Discount = Val(lblHeading.Caption)
            End If
            Unload Me
        Case "CL"
            lblHeading = "0.00"
            cmdErr.Visible = False
            errTimer.Enabled = False
    End Select
End Sub
Private Sub cmdKey_Click(Index As Integer)
    TillData.Discount = 0
    Select Case cmdKey(Index).Caption
        Case "-10%": TillData.Discount = 10
        Case "-20%": TillData.Discount = 20
        Case "-30%": TillData.Discount = 30
        Case "-40%": TillData.Discount = 40
        Case "-50%": TillData.Discount = 50
        Case "-60%": TillData.Discount = 60
        Case "-70%": TillData.Discount = 70
        Case "-80%": TillData.Discount = 80
        Case "Cost +10": TillData.Discount = -10
        Case "Free": TillData.Discount = 100
    End Select
    Unload Me
End Sub
Private Sub errTimer_Timer()
    Select Case cmdErr.BackColor
        Case &HC0C0&          'White
            cmdErr.BackColor = &HFF&
        Case &HFF&             'Yellow
            cmdErr.BackColor = &HC0C0&
    End Select
End Sub
Private Sub Form_Activate()
    If Workstation.Disc10 = 1 Then cmdKey(0).Enabled = True Else cmdKey(0).Enabled = False
    If Workstation.Disc20 = 1 Then cmdKey(1).Enabled = True Else cmdKey(1).Enabled = False
    If Workstation.Disc30 = 1 Then cmdKey(2).Enabled = True Else cmdKey(2).Enabled = False
    If Workstation.Disc40 = 1 Then cmdKey(3).Enabled = True Else cmdKey(3).Enabled = False
    If Workstation.Disc50 = 1 Then cmdKey(4).Enabled = True Else cmdKey(4).Enabled = False
    If Workstation.Disc60 = 1 Then cmdKey(5).Enabled = True Else cmdKey(5).Enabled = False
    If Workstation.Disc70 = 1 Then cmdKey(6).Enabled = True Else cmdKey(6).Enabled = False
    If Workstation.Disc80 = 1 Then cmdKey(7).Enabled = True Else cmdKey(7).Enabled = False
    If Workstation.Disc90 = 1 Then cmdKey(8).Enabled = True Else cmdKey(8).Enabled = False
    If Workstation.DiscFree = 1 Then cmdKey(9).Enabled = True Else cmdKey(9).Enabled = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case cmdApply.Caption
        Case "Whole Sale": TillData.ExtraFunc = "Sale Discount"
        Case "Current Line Only": TillData.ExtraFunc = "Line Discount"
    End Select
End Sub
