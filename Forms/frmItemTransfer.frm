VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmItemTransfer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11520
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmItemTransfer.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1515
      Index           =   16
      Left            =   10290
      TabIndex        =   28
      Top             =   9075
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2672
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":D21D
      textLT          =   "frmItemTransfer.frx":D235
      textCT          =   "frmItemTransfer.frx":D24D
      textRT          =   "frmItemTransfer.frx":D265
      textLM          =   "frmItemTransfer.frx":D27D
      textRM          =   "frmItemTransfer.frx":D295
      textLB          =   "frmItemTransfer.frx":D2AD
      textCB          =   "frmItemTransfer.frx":D2C5
      textRB          =   "frmItemTransfer.frx":D2DD
      colorBack       =   "frmItemTransfer.frx":D2F5
      colorIntern     =   "frmItemTransfer.frx":D31F
      colorMO         =   "frmItemTransfer.frx":D349
      colorFocus      =   "frmItemTransfer.frx":D373
      colorDisabled   =   "frmItemTransfer.frx":D39D
      colorPressed    =   "frmItemTransfer.frx":D3C7
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1515
      Index           =   17
      Left            =   12690
      TabIndex        =   29
      Top             =   9075
      Visible         =   0   'False
      Width           =   2400
      _Version        =   524298
      _ExtentX        =   4233
      _ExtentY        =   2672
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":D3F1
      textLT          =   "frmItemTransfer.frx":D409
      textCT          =   "frmItemTransfer.frx":D421
      textRT          =   "frmItemTransfer.frx":D439
      textLM          =   "frmItemTransfer.frx":D451
      textRM          =   "frmItemTransfer.frx":D469
      textLB          =   "frmItemTransfer.frx":D481
      textCB          =   "frmItemTransfer.frx":D499
      textRB          =   "frmItemTransfer.frx":D4B1
      colorBack       =   "frmItemTransfer.frx":D4C9
      colorIntern     =   "frmItemTransfer.frx":D4F3
      colorMO         =   "frmItemTransfer.frx":D51D
      colorFocus      =   "frmItemTransfer.frx":D547
      colorDisabled   =   "frmItemTransfer.frx":D571
      colorPressed    =   "frmItemTransfer.frx":D59B
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1545
      Index           =   13
      Left            =   10290
      TabIndex        =   25
      Top             =   7590
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2725
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":D5C5
      textLT          =   "frmItemTransfer.frx":D5DD
      textCT          =   "frmItemTransfer.frx":D5F5
      textRT          =   "frmItemTransfer.frx":D60D
      textLM          =   "frmItemTransfer.frx":D625
      textRM          =   "frmItemTransfer.frx":D63D
      textLB          =   "frmItemTransfer.frx":D655
      textCB          =   "frmItemTransfer.frx":D66D
      textRB          =   "frmItemTransfer.frx":D685
      colorBack       =   "frmItemTransfer.frx":D69D
      colorIntern     =   "frmItemTransfer.frx":D6C7
      colorMO         =   "frmItemTransfer.frx":D6F1
      colorFocus      =   "frmItemTransfer.frx":D71B
      colorDisabled   =   "frmItemTransfer.frx":D745
      colorPressed    =   "frmItemTransfer.frx":D76F
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1545
      Index           =   14
      Left            =   12690
      TabIndex        =   26
      Top             =   7590
      Visible         =   0   'False
      Width           =   2400
      _Version        =   524298
      _ExtentX        =   4233
      _ExtentY        =   2725
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":D799
      textLT          =   "frmItemTransfer.frx":D7B1
      textCT          =   "frmItemTransfer.frx":D7C9
      textRT          =   "frmItemTransfer.frx":D7E1
      textLM          =   "frmItemTransfer.frx":D7F9
      textRM          =   "frmItemTransfer.frx":D811
      textLB          =   "frmItemTransfer.frx":D829
      textCB          =   "frmItemTransfer.frx":D841
      textRB          =   "frmItemTransfer.frx":D859
      colorBack       =   "frmItemTransfer.frx":D871
      colorIntern     =   "frmItemTransfer.frx":D89B
      colorMO         =   "frmItemTransfer.frx":D8C5
      colorFocus      =   "frmItemTransfer.frx":D8EF
      colorDisabled   =   "frmItemTransfer.frx":D919
      colorPressed    =   "frmItemTransfer.frx":D943
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   10
      Left            =   10290
      TabIndex        =   22
      Top             =   6045
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":D96D
      textLT          =   "frmItemTransfer.frx":D985
      textCT          =   "frmItemTransfer.frx":D99D
      textRT          =   "frmItemTransfer.frx":D9B5
      textLM          =   "frmItemTransfer.frx":D9CD
      textRM          =   "frmItemTransfer.frx":D9E5
      textLB          =   "frmItemTransfer.frx":D9FD
      textCB          =   "frmItemTransfer.frx":DA15
      textRB          =   "frmItemTransfer.frx":DA2D
      colorBack       =   "frmItemTransfer.frx":DA45
      colorIntern     =   "frmItemTransfer.frx":DA6F
      colorMO         =   "frmItemTransfer.frx":DA99
      colorFocus      =   "frmItemTransfer.frx":DAC3
      colorDisabled   =   "frmItemTransfer.frx":DAED
      colorPressed    =   "frmItemTransfer.frx":DB17
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   11
      Left            =   12690
      TabIndex        =   23
      Top             =   6045
      Visible         =   0   'False
      Width           =   2400
      _Version        =   524298
      _ExtentX        =   4233
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":DB41
      textLT          =   "frmItemTransfer.frx":DB59
      textCT          =   "frmItemTransfer.frx":DB71
      textRT          =   "frmItemTransfer.frx":DB89
      textLM          =   "frmItemTransfer.frx":DBA1
      textRM          =   "frmItemTransfer.frx":DBB9
      textLB          =   "frmItemTransfer.frx":DBD1
      textCB          =   "frmItemTransfer.frx":DBE9
      textRB          =   "frmItemTransfer.frx":DC01
      colorBack       =   "frmItemTransfer.frx":DC19
      colorIntern     =   "frmItemTransfer.frx":DC43
      colorMO         =   "frmItemTransfer.frx":DC6D
      colorFocus      =   "frmItemTransfer.frx":DC97
      colorDisabled   =   "frmItemTransfer.frx":DCC1
      colorPressed    =   "frmItemTransfer.frx":DCEB
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   7
      Left            =   10290
      TabIndex        =   19
      Top             =   4500
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":DD15
      textLT          =   "frmItemTransfer.frx":DD2D
      textCT          =   "frmItemTransfer.frx":DD45
      textRT          =   "frmItemTransfer.frx":DD5D
      textLM          =   "frmItemTransfer.frx":DD75
      textRM          =   "frmItemTransfer.frx":DD8D
      textLB          =   "frmItemTransfer.frx":DDA5
      textCB          =   "frmItemTransfer.frx":DDBD
      textRB          =   "frmItemTransfer.frx":DDD5
      colorBack       =   "frmItemTransfer.frx":DDED
      colorIntern     =   "frmItemTransfer.frx":DE17
      colorMO         =   "frmItemTransfer.frx":DE41
      colorFocus      =   "frmItemTransfer.frx":DE6B
      colorDisabled   =   "frmItemTransfer.frx":DE95
      colorPressed    =   "frmItemTransfer.frx":DEBF
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   8
      Left            =   12690
      TabIndex        =   20
      Top             =   4500
      Visible         =   0   'False
      Width           =   2400
      _Version        =   524298
      _ExtentX        =   4233
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":DEE9
      textLT          =   "frmItemTransfer.frx":DF01
      textCT          =   "frmItemTransfer.frx":DF19
      textRT          =   "frmItemTransfer.frx":DF31
      textLM          =   "frmItemTransfer.frx":DF49
      textRM          =   "frmItemTransfer.frx":DF61
      textLB          =   "frmItemTransfer.frx":DF79
      textCB          =   "frmItemTransfer.frx":DF91
      textRB          =   "frmItemTransfer.frx":DFA9
      colorBack       =   "frmItemTransfer.frx":DFC1
      colorIntern     =   "frmItemTransfer.frx":DFEB
      colorMO         =   "frmItemTransfer.frx":E015
      colorFocus      =   "frmItemTransfer.frx":E03F
      colorDisabled   =   "frmItemTransfer.frx":E069
      colorPressed    =   "frmItemTransfer.frx":E093
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   5
      Left            =   12690
      TabIndex        =   17
      Top             =   2955
      Visible         =   0   'False
      Width           =   2400
      _Version        =   524298
      _ExtentX        =   4233
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":E0BD
      textLT          =   "frmItemTransfer.frx":E0D5
      textCT          =   "frmItemTransfer.frx":E0ED
      textRT          =   "frmItemTransfer.frx":E105
      textLM          =   "frmItemTransfer.frx":E11D
      textRM          =   "frmItemTransfer.frx":E135
      textLB          =   "frmItemTransfer.frx":E14D
      textCB          =   "frmItemTransfer.frx":E165
      textRB          =   "frmItemTransfer.frx":E17D
      colorBack       =   "frmItemTransfer.frx":E195
      colorIntern     =   "frmItemTransfer.frx":E1BF
      colorMO         =   "frmItemTransfer.frx":E1E9
      colorFocus      =   "frmItemTransfer.frx":E213
      colorDisabled   =   "frmItemTransfer.frx":E23D
      colorPressed    =   "frmItemTransfer.frx":E267
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   2
      Left            =   12690
      TabIndex        =   14
      Top             =   1410
      Visible         =   0   'False
      Width           =   2400
      _Version        =   524298
      _ExtentX        =   4233
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":E291
      textLT          =   "frmItemTransfer.frx":E2A9
      textCT          =   "frmItemTransfer.frx":E2C1
      textRT          =   "frmItemTransfer.frx":E2D9
      textLM          =   "frmItemTransfer.frx":E2F1
      textRM          =   "frmItemTransfer.frx":E309
      textLB          =   "frmItemTransfer.frx":E321
      textCB          =   "frmItemTransfer.frx":E339
      textRB          =   "frmItemTransfer.frx":E351
      colorBack       =   "frmItemTransfer.frx":E369
      colorIntern     =   "frmItemTransfer.frx":E393
      colorMO         =   "frmItemTransfer.frx":E3BD
      colorFocus      =   "frmItemTransfer.frx":E3E7
      colorDisabled   =   "frmItemTransfer.frx":E411
      colorPressed    =   "frmItemTransfer.frx":E43B
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9195
      Left            =   330
      ScaleHeight     =   9165
      ScaleWidth      =   6105
      TabIndex        =   1
      Top             =   1380
      Width           =   6135
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   675
         Index           =   0
         Left            =   5355
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   1191
         Appearance      =   3
         BackColor       =   7848417
         BorderColor     =   8421504
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
      Begin VSFlex8Ctl.VSFlexGrid grdMain 
         Height          =   9150
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5355
         _cx             =   9446
         _cy             =   16140
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   20
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   660
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
         WallPaper       =   "frmItemTransfer.frx":E465
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   630
         Index           =   1
         Left            =   5355
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   8550
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   1111
         Appearance      =   3
         BackColor       =   7848417
         BorderColor     =   8421504
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
      Begin VB.Shape picGrid 
         BackColor       =   &H00B4DAED&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0086C4E1&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0086C4E1&
         FillStyle       =   2  'Horizontal Line
         Height          =   8505
         Left            =   5370
         Top             =   510
         Width           =   735
      End
   End
   Begin VB.Timer scrolTimer 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   630
   End
   Begin VB.Timer scrolTimer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   780
      Top             =   90
   End
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   930
      Left            =   14130
      TabIndex        =   5
      Top             =   330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1640
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
      ShowFocus       =   0
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1620
      Index           =   0
      Left            =   6570
      TabIndex        =   0
      Top             =   1410
      Width           =   1230
      _Version        =   524298
      _ExtentX        =   2170
      _ExtentY        =   2857
      _StockProps     =   66
      Caption         =   "To Table"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1744599
      LogPixels       =   96
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":12196
      textLT          =   "frmItemTransfer.frx":12206
      textCT          =   "frmItemTransfer.frx":1221E
      textRT          =   "frmItemTransfer.frx":12236
      textLM          =   "frmItemTransfer.frx":1224E
      textRM          =   "frmItemTransfer.frx":12266
      textLB          =   "frmItemTransfer.frx":1227E
      textCB          =   "frmItemTransfer.frx":12296
      textRB          =   "frmItemTransfer.frx":122AE
      colorBack       =   "frmItemTransfer.frx":122C6
      colorIntern     =   "frmItemTransfer.frx":122F0
      colorMO         =   "frmItemTransfer.frx":1231A
      colorFocus      =   "frmItemTransfer.frx":12344
      colorDisabled   =   "frmItemTransfer.frx":1236E
      colorPressed    =   "frmItemTransfer.frx":12398
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1590
      Index           =   1
      Left            =   6570
      TabIndex        =   10
      Top             =   3180
      Width           =   1230
      _Version        =   524298
      _ExtentX        =   2170
      _ExtentY        =   2805
      _StockProps     =   66
      Caption         =   "To Tab"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1744599
      LogPixels       =   96
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":123C2
      textLT          =   "frmItemTransfer.frx":1242E
      textCT          =   "frmItemTransfer.frx":12446
      textRT          =   "frmItemTransfer.frx":1245E
      textLM          =   "frmItemTransfer.frx":12476
      textRM          =   "frmItemTransfer.frx":1248E
      textLB          =   "frmItemTransfer.frx":124A6
      textCB          =   "frmItemTransfer.frx":124BE
      textRB          =   "frmItemTransfer.frx":124D6
      colorBack       =   "frmItemTransfer.frx":124EE
      colorIntern     =   "frmItemTransfer.frx":12518
      colorMO         =   "frmItemTransfer.frx":12542
      colorFocus      =   "frmItemTransfer.frx":1256C
      colorDisabled   =   "frmItemTransfer.frx":12596
      colorPressed    =   "frmItemTransfer.frx":125C0
      Style           =   2
      Orientation     =   4
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   150
      Index           =   3
      Left            =   6570
      TabIndex        =   11
      Top             =   3030
      Width           =   1230
      _Version        =   524298
      _ExtentX        =   2170
      _ExtentY        =   265
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1744599
      LogPixels       =   96
      Clickable       =   0   'False
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":125EA
      textLT          =   "frmItemTransfer.frx":12602
      textCT          =   "frmItemTransfer.frx":1261A
      textRT          =   "frmItemTransfer.frx":12632
      textLM          =   "frmItemTransfer.frx":1264A
      textRM          =   "frmItemTransfer.frx":12662
      textLB          =   "frmItemTransfer.frx":1267A
      textCB          =   "frmItemTransfer.frx":12692
      textRB          =   "frmItemTransfer.frx":126AA
      colorBack       =   "frmItemTransfer.frx":126C2
      colorIntern     =   "frmItemTransfer.frx":126EC
      colorMO         =   "frmItemTransfer.frx":12716
      colorFocus      =   "frmItemTransfer.frx":12740
      colorDisabled   =   "frmItemTransfer.frx":1276A
      colorPressed    =   "frmItemTransfer.frx":12794
      Orientation     =   4
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   1
      Left            =   10290
      TabIndex        =   13
      Top             =   1410
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":127BE
      textLT          =   "frmItemTransfer.frx":127D6
      textCT          =   "frmItemTransfer.frx":127EE
      textRT          =   "frmItemTransfer.frx":12806
      textLM          =   "frmItemTransfer.frx":1281E
      textRM          =   "frmItemTransfer.frx":12836
      textLB          =   "frmItemTransfer.frx":1284E
      textCB          =   "frmItemTransfer.frx":12866
      textRB          =   "frmItemTransfer.frx":1287E
      colorBack       =   "frmItemTransfer.frx":12896
      colorIntern     =   "frmItemTransfer.frx":128C0
      colorMO         =   "frmItemTransfer.frx":128EA
      colorFocus      =   "frmItemTransfer.frx":12914
      colorDisabled   =   "frmItemTransfer.frx":1293E
      colorPressed    =   "frmItemTransfer.frx":12968
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   3
      Left            =   7890
      TabIndex        =   15
      Top             =   2955
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":12992
      textLT          =   "frmItemTransfer.frx":129AA
      textCT          =   "frmItemTransfer.frx":129C2
      textRT          =   "frmItemTransfer.frx":129DA
      textLM          =   "frmItemTransfer.frx":129F2
      textRM          =   "frmItemTransfer.frx":12A0A
      textLB          =   "frmItemTransfer.frx":12A22
      textCB          =   "frmItemTransfer.frx":12A3A
      textRB          =   "frmItemTransfer.frx":12A52
      colorBack       =   "frmItemTransfer.frx":12A6A
      colorIntern     =   "frmItemTransfer.frx":12A94
      colorMO         =   "frmItemTransfer.frx":12ABE
      colorFocus      =   "frmItemTransfer.frx":12AE8
      colorDisabled   =   "frmItemTransfer.frx":12B12
      colorPressed    =   "frmItemTransfer.frx":12B3C
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   4
      Left            =   10290
      TabIndex        =   16
      Top             =   2955
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":12B66
      textLT          =   "frmItemTransfer.frx":12B7E
      textCT          =   "frmItemTransfer.frx":12B96
      textRT          =   "frmItemTransfer.frx":12BAE
      textLM          =   "frmItemTransfer.frx":12BC6
      textRM          =   "frmItemTransfer.frx":12BDE
      textLB          =   "frmItemTransfer.frx":12BF6
      textCB          =   "frmItemTransfer.frx":12C0E
      textRB          =   "frmItemTransfer.frx":12C26
      colorBack       =   "frmItemTransfer.frx":12C3E
      colorIntern     =   "frmItemTransfer.frx":12C68
      colorMO         =   "frmItemTransfer.frx":12C92
      colorFocus      =   "frmItemTransfer.frx":12CBC
      colorDisabled   =   "frmItemTransfer.frx":12CE6
      colorPressed    =   "frmItemTransfer.frx":12D10
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdTable 
      Height          =   7890
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   75
      _cx             =   132
      _cy             =   13917
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
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   200
      ColWidthMax     =   200
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
   Begin VSFlex8Ctl.VSFlexGrid grdTab 
      Height          =   7890
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   75
      _cx             =   132
      _cy             =   13917
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
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   200
      ColWidthMax     =   200
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
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   0
      Left            =   7890
      TabIndex        =   12
      Top             =   1410
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":12D3A
      textLT          =   "frmItemTransfer.frx":12D52
      textCT          =   "frmItemTransfer.frx":12D6A
      textRT          =   "frmItemTransfer.frx":12D82
      textLM          =   "frmItemTransfer.frx":12D9A
      textRM          =   "frmItemTransfer.frx":12DB2
      textLB          =   "frmItemTransfer.frx":12DCA
      textCB          =   "frmItemTransfer.frx":12DE2
      textRB          =   "frmItemTransfer.frx":12DFA
      colorBack       =   "frmItemTransfer.frx":12E12
      colorIntern     =   "frmItemTransfer.frx":12E3C
      colorMO         =   "frmItemTransfer.frx":12E66
      colorFocus      =   "frmItemTransfer.frx":12E90
      colorDisabled   =   "frmItemTransfer.frx":12EBA
      colorPressed    =   "frmItemTransfer.frx":12EE4
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   6
      Left            =   7890
      TabIndex        =   18
      Top             =   4500
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":12F0E
      textLT          =   "frmItemTransfer.frx":12F26
      textCT          =   "frmItemTransfer.frx":12F3E
      textRT          =   "frmItemTransfer.frx":12F56
      textLM          =   "frmItemTransfer.frx":12F6E
      textRM          =   "frmItemTransfer.frx":12F86
      textLB          =   "frmItemTransfer.frx":12F9E
      textCB          =   "frmItemTransfer.frx":12FB6
      textRB          =   "frmItemTransfer.frx":12FCE
      colorBack       =   "frmItemTransfer.frx":12FE6
      colorIntern     =   "frmItemTransfer.frx":13010
      colorMO         =   "frmItemTransfer.frx":1303A
      colorFocus      =   "frmItemTransfer.frx":13064
      colorDisabled   =   "frmItemTransfer.frx":1308E
      colorPressed    =   "frmItemTransfer.frx":130B8
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1575
      Index           =   9
      Left            =   7890
      TabIndex        =   21
      Top             =   6045
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2778
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":130E2
      textLT          =   "frmItemTransfer.frx":130FA
      textCT          =   "frmItemTransfer.frx":13112
      textRT          =   "frmItemTransfer.frx":1312A
      textLM          =   "frmItemTransfer.frx":13142
      textRM          =   "frmItemTransfer.frx":1315A
      textLB          =   "frmItemTransfer.frx":13172
      textCB          =   "frmItemTransfer.frx":1318A
      textRB          =   "frmItemTransfer.frx":131A2
      colorBack       =   "frmItemTransfer.frx":131BA
      colorIntern     =   "frmItemTransfer.frx":131E4
      colorMO         =   "frmItemTransfer.frx":1320E
      colorFocus      =   "frmItemTransfer.frx":13238
      colorDisabled   =   "frmItemTransfer.frx":13262
      colorPressed    =   "frmItemTransfer.frx":1328C
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1545
      Index           =   12
      Left            =   7890
      TabIndex        =   24
      Top             =   7590
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2725
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":132B6
      textLT          =   "frmItemTransfer.frx":132CE
      textCT          =   "frmItemTransfer.frx":132E6
      textRT          =   "frmItemTransfer.frx":132FE
      textLM          =   "frmItemTransfer.frx":13316
      textRM          =   "frmItemTransfer.frx":1332E
      textLB          =   "frmItemTransfer.frx":13346
      textCB          =   "frmItemTransfer.frx":1335E
      textRB          =   "frmItemTransfer.frx":13376
      colorBack       =   "frmItemTransfer.frx":1338E
      colorIntern     =   "frmItemTransfer.frx":133B8
      colorMO         =   "frmItemTransfer.frx":133E2
      colorFocus      =   "frmItemTransfer.frx":1340C
      colorDisabled   =   "frmItemTransfer.frx":13436
      colorPressed    =   "frmItemTransfer.frx":13460
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1515
      Index           =   15
      Left            =   7890
      TabIndex        =   27
      Top             =   9075
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2672
      _StockProps     =   66
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
         Name            =   "Arial Narrow"
         Size            =   12
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
      BackColorContainer=   4241128
      SpecialEffect   =   1
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmItemTransfer.frx":1348A
      textLT          =   "frmItemTransfer.frx":134A2
      textCT          =   "frmItemTransfer.frx":134BA
      textRT          =   "frmItemTransfer.frx":134D2
      textLM          =   "frmItemTransfer.frx":134EA
      textRM          =   "frmItemTransfer.frx":13502
      textLB          =   "frmItemTransfer.frx":1351A
      textCB          =   "frmItemTransfer.frx":13532
      textRB          =   "frmItemTransfer.frx":1354A
      colorBack       =   "frmItemTransfer.frx":13562
      colorIntern     =   "frmItemTransfer.frx":1358C
      colorMO         =   "frmItemTransfer.frx":135B6
      colorFocus      =   "frmItemTransfer.frx":135E0
      colorDisabled   =   "frmItemTransfer.frx":1360A
      colorPressed    =   "frmItemTransfer.frx":13634
      Orientation     =   8
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin MSForms.Label lblUser 
      Height          =   285
      Left            =   10440
      TabIndex        =   9
      Top             =   10860
      Width           =   4335
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7646;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   540
      TabIndex        =   8
      Top             =   10860
      Width           =   4365
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7699;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblKeyRegister 
      Height          =   675
      Left            =   660
      TabIndex        =   7
      Top             =   570
      Width           =   6795
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "11986;1191"
      FontName        =   "Arial Narrow"
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblTotal 
      Height          =   675
      Left            =   7380
      TabIndex        =   6
      Top             =   570
      Width           =   6315
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "11139;1191"
      FontName        =   "Arial Narrow"
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
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
      Picture         =   "frmItemTransfer.frx":1365E
   End
End
Attribute VB_Name = "frmItemTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdArrow_Click(Index As Integer)
    Select Case Index
        Case 0
            If grdMain(0).Row <> 1 Then
                grdMain(0).Row = grdMain(0).Row - 1
            End If
        Case 1
            If grdMain(0).Row <> grdMain(0).Rows - 1 Then
                grdMain(0).Row = grdMain(0).Row + 1
            End If
    End Select
    grdMain(0).ShowCell grdMain(0).Row, 0
End Sub
Private Sub cmdClose_Click()
    Update_Sale
    Unload Me
End Sub
Private Sub cmdDeptStrip_Click(Index As Integer)
    Select Case Index
        Case 0
            LoadTables 0
        Case 1
            LoadTabs 0
    End Select
End Sub

Private Sub cmdTable_Click(Index As Integer)
If cmdDeptStrip(0).Value = 1 Then
    DoEvents
    If cmdTable(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        grdTable.Row = grdTable.Row + 1
        For i = 0 To 17
            
            
            
            
            
            If grdTable.TextMatrix(grdTable.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                Else
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                    grdTable.Row = grdTable.Row - 1
                    Exit For
                End If
            Else
                cmdTable(i).Caption = "Table No: " & grdTable.TextMatrix(grdTable.Row, 0)
                cmdTable(i).Tag = grdTable.TextMatrix(grdTable.Row, 1)
                cmdTable(i).TextDescrCB.OffsetY = -10
                cmdTable(i).TextDescrCB.ColorNormal = &H800000
                cmdTable(i).TextDescrCB.Text = grdTable.TextMatrix(grdTable.Row, 2)
                If grdTable.TextMatrix(grdTable.Row, 3) = "True" And Workstation_No <> Val(grdTable.TextMatrix(grdTable.Row, 4)) Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTable(i).TextDescrCT.Text = ""
                End If
                If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
            End If
            If grdTable.Row = grdTable.Rows - 1 Then Exit For
            grdTable.Row = grdTable.Row + 1
        Next i
        For b = i + 1 To cmdTable.Count - 1
            cmdTable(b).Caption = "1"
            cmdTable(b).Tag = ""
            cmdTable(b).Visible = False
        Next b
        Exit Sub
    End If
    If cmdTable(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdTable(0).Picture = ""
        While grdTable.TextMatrix(grdTable.Row, 0) <> "ßArrow"
            grdTable.Row = grdTable.Row - 1
        Wend
        grdTable.Row = grdTable.Row - 17
        For i = 0 To 17
            If grdTable.TextMatrix(grdTable.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                Else
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                    grdTable.Row = grdTable.Row - 1
                    Exit For
                End If
            Else
                cmdTable(i).Caption = "Table No: " & grdTable.TextMatrix(grdTable.Row, 0)
                cmdTable(i).Tag = grdTable.TextMatrix(grdTable.Row, 1)
                cmdTable(i).TextDescrCB.OffsetY = -10
                cmdTable(i).TextDescrCB.ColorNormal = &H800000
                cmdTable(i).TextDescrCB.Text = grdTable.TextMatrix(grdTable.Row, 2)
                If grdTable.TextMatrix(grdTable.Row, 3) = "True" And Workstation_No <> Val(grdTable.TextMatrix(grdTable.Row, 4)) Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTable(i).TextDescrCT.Text = ""
                End If
                If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
            End If
            If grdTable.Row = grdTable.Rows - 1 Then Exit For
            grdTable.Row = grdTable.Row + 1
        Next i
        For b = i + 1 To cmdTable.Count - 1
            cmdTable(b).Caption = "1"
            cmdTable(b).Tag = ""
            cmdTable(b).TextDescrCB.Text = ""
            cmdTable(b).TextDescrCT.Text = ""
            cmdTable(b).ToolTipText = ""
            cmdTable(b).Visible = False
        Next b
        Exit Sub
    End If
End If
If cmdDeptStrip(1).Value = 1 Then
    If cmdTable(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        grdTab.Row = grdTab.Row + 1
        For i = 0 To 17
            If grdTab.TextMatrix(grdTab.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                Else
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                    grdTab.Row = grdTab.Row - 1
                    Exit For
                End If
            Else
                cmdTable(i).Caption = grdTab.TextMatrix(grdTab.Row, 6)
                cmdTable(i).Tag = grdTab.TextMatrix(grdTab.Row, 0)
                cmdTable(i).TextDescrCB.OffsetY = -10
                cmdTable(i).TextDescrCB.ColorNormal = &H800000
                cmdTable(i).TextDescrCB.Text = grdTab.TextMatrix(grdTab.Row, 2)
                If grdTab.TextMatrix(grdTab.Row, 3) = "True" And Workstation_No <> Val(grdTab.TextMatrix(grdTab.Row, 4)) Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    If Trim(grdTab.TextMatrix(grdTab.Row, 5)) = "" Then
                        'cmdTable(i).TextDescrCB.Text = ""
                    Else
                        cmdTable(i).TextDescrCB.OffsetY = -10
                        cmdTable(i).TextDescrCB.ColorNormal = &H800000
                        cmdTable(i).TextDescrCB.Text = "From > " & grdTab.TextMatrix(grdTab.Row, 5)
                    End If
                End If
            End If
            If grdTab.Row = grdTab.Rows - 1 Then Exit For
            grdTab.Row = grdTab.Row + 1
        Next i
        For b = i + 1 To cmdTable.Count - 1
            cmdTable(b).Caption = "1"
            cmdTable(b).Tag = ""
            cmdTable(b).Visible = False
        Next b
        Exit Sub
    End If
    If cmdTable(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdTable(0).Picture = ""
        While grdTab.TextMatrix(grdTab.Row, 0) <> "ßArrow"
            grdTab.Row = grdTab.Row - 1
        Wend
        grdTab.Row = grdTab.Row - 17
        For i = 0 To 17
            If grdTab.TextMatrix(grdTab.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                Else
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                    grdTab.Row = grdTab.Row - 1
                    Exit For
                End If
            Else
                cmdTable(i).Caption = grdTab.TextMatrix(grdTab.Row, 6)
                cmdTable(i).Tag = grdTab.TextMatrix(grdTab.Row, 0)
                cmdTable(i).TextDescrCB.OffsetY = -10
                cmdTable(i).TextDescrCB.ColorNormal = &H800000
                cmdTable(i).TextDescrCB.Text = grdTab.TextMatrix(grdTab.Row, 2)
                If grdTab.TextMatrix(grdTab.Row, 3) = "True" And Workstation_No <> Val(grdTab.TextMatrix(grdTab.Row, 4)) Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTable(i).TextDescrCT.Text = ""
                End If
                If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
            End If
            If grdTab.Row = grdTab.Rows - 1 Then Exit For
            grdTab.Row = grdTab.Row + 1
        Next i
        For b = i + 1 To cmdTable.Count - 1
            cmdTable(b).Caption = "1"
            cmdTable(b).Tag = ""
            cmdTable(b).TextDescrCB.Text = ""
            cmdTable(b).TextDescrCT.Text = ""
            cmdTable(b).ToolTipText = ""
            cmdTable(b).Visible = False
        Next b
        Exit Sub
    End If
End If
If cmdDeptStrip(0).Value = 1 Then
top:
    For i = 1 To grdMain(0).Rows - 1
        If grdMain(0).TextMatrix(i, 18) = "1" Then
            ActiveReadServer "Select Workstation_No,Doc_No, Table_name from Table_Listing where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStrRev(cmdTable(Index).Caption, ":") + 1))
            Workstation1 = rs.Fields("Workstation_No")
            Doc_No = rs.Fields("Doc_No")
            Table_Name = rs.Fields("Table_name")
            rs.Close
            ActiveUpdateServer "INSERT INTO [Table_Listing]([Table_No],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value, Table_name)" & _
            " VALUES('" & Val(Mid(cmdTable(Index).Caption, InStrRev(cmdTable(Index).Caption, ":") + 1)) & "','" & Val(cmdTable(Index).Tag) & "','" & Val(Mid(cmdTable(Index).TextDescrCB.Text, 1, InStr(cmdTable(Index).TextDescrCB.Text, "-") - 1)) & "','" & Workstation1 & "','" & grdMain(0).TextMatrix(i, 0) & "','" & grdMain(0).TextMatrix(i, 1) & "','" & _
            grdMain(0).TextMatrix(i, 2) & "','" & grdMain(0).TextMatrix(i, 3) & "','" & grdMain(0).TextMatrix(i, 4) & "','" & grdMain(0).TextMatrix(i, 5) & "','" & grdMain(0).TextMatrix(i, 6) & "','" & grdMain(0).TextMatrix(i, 8) & "','" & grdMain(0).TextMatrix(i, 9) & "','" & grdMain(0).TextMatrix(i, 10) & "','" & grdMain(0).TextMatrix(i, 11) & "','" & grdMain(0).TextMatrix(i, 12) & "','" & grdMain(0).TextMatrix(i, 13) & "','" & grdMain(0).TextMatrix(i, 14) & "','" & grdMain(0).TextMatrix(i, 7) & "'," & Doc_No & ",0,'" & grdMain(0).TextMatrix(i, 17) & "',0," & grdMain(0).ValueMatrix(i, 19) & ",'" & Table_Name & "')"
            grdMain(0).RemoveItem (i)
            GoTo top
        End If
    Next i
    Update_Sale
End If
If cmdDeptStrip(1).Value = 1 Then
top1:
    For i = 1 To grdMain(0).Rows - 1
        If grdMain(0).TextMatrix(i, 18) = "1" Then
            ActiveReadServer "Select Workstation_No,Doc_No,Tab_Name from Tab_Listing where Tab_No = " & cmdTable(Index).Tag
            Workstation1 = rs.Fields("Workstation_No")
            Doc_No = rs.Fields("Doc_No")
            Tab_Name = rs.Fields("Tab_Name")
            rs.Close
            ActiveUpdateServer "INSERT INTO [Tab_Listing]([Tab_No],[Tab_Name],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value)" & _
            " VALUES('" & cmdTable(Index).Tag & "','" & Tab_Name & "','0','" & Val(Mid(cmdTable(Index).TextDescrCB.Text, 1, InStr(cmdTable(Index).TextDescrCB.Text, "-") - 1)) & "','" & Workstation1 & "','" & grdMain(0).TextMatrix(i, 0) & "','" & grdMain(0).TextMatrix(i, 1) & "','" & _
            grdMain(0).TextMatrix(i, 2) & "','" & grdMain(0).TextMatrix(i, 3) & "','" & grdMain(0).TextMatrix(i, 4) & "','" & grdMain(0).TextMatrix(i, 5) & "','" & grdMain(0).TextMatrix(i, 6) & "','" & grdMain(0).TextMatrix(i, 8) & "','" & grdMain(0).TextMatrix(i, 9) & "','" & grdMain(0).TextMatrix(i, 10) & "','" & grdMain(0).TextMatrix(i, 11) & "','" & grdMain(0).TextMatrix(i, 12) & "','" & grdMain(0).TextMatrix(i, 13) & "','" & grdMain(0).TextMatrix(i, 14) & "','" & grdMain(0).TextMatrix(i, 7) & "'," & Doc_No & ",0,'" & grdMain(0).TextMatrix(i, 17) & "',0," & grdMain(0).ValueMatrix(i, 19) & ")"
            grdMain(0).RemoveItem (i)
            GoTo top1
        End If
    Next i
    Update_Sale
End If
Unload Me
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 1
    If TillData.TableNo <> 0 Then
        cmdDeptStrip(0).Value = -1
        LoadTables 0
    End If
    If TillData.TabNo <> 0 Then
        cmdDeptStrip(1).Value = -1
        LoadTabs 0
    End If
    grdMain(0).Rows = 1
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
    grdMain(0).Rows = 1
    grdMain(0).SetFocus
    Select Case Panel_no
        Case 1
            With frmSales1
                For i = 1 To .grdMain.Rows - 1
                     If .grdMain.TextMatrix(i, 0) <> "" Then
                        If .grdMain.TextMatrix(i, 8) = "" Or .grdMain.TextMatrix(i, 8) = "Return Item" Then
                            grdMain(0).Rows = grdMain(0).Rows + 1
                            For b = 0 To .grdMain.Cols - 1
                               grdMain(0).TextMatrix(grdMain(0).Rows - 1, b) = .grdMain.TextMatrix(i, b)
                               If b = 18 Then grdMain(0).TextMatrix(grdMain(0).Rows - 1, b) = "0"
                            Next b
                        End If
                    End If
                Next i
                grdMain(0).Row = 1
                grdMain(0).ShowCell 1, 0
            End With
        Case 2
            With frmBar
                For i = 1 To .grdMain.Rows - 1
                     If .grdMain.TextMatrix(i, 0) <> "" Then
                         If .grdMain.TextMatrix(i, 8) = "" Or .grdMain.TextMatrix(i, 8) = "Return Item" Then
                            grdMain(0).Rows = grdMain(0).Rows + 1
                            For b = 0 To .grdMain.Cols - 1
                               grdMain(0).TextMatrix(grdMain(0).Rows - 1, b) = .grdMain.TextMatrix(i, b)
                               If b = 18 Then grdMain(0).TextMatrix(grdMain(0).Rows - 1, b) = "0"
                            Next b
                        End If
                    End If
                Next i
                grdMain(0).Row = 1
                grdMain(0).ShowCell 1, 0
            End With
    End Select
    Select Case Panel_no
        Case 1
            lblTotal.Caption = "Subtotal: " & frmSales1.lblTender
        Case 2
            lblTotal.Caption = "Subtotal: " & frmBar.lblTender
    End Select
End Sub
Private Sub Form_Load()
    grdMain(i).Rows = 1
    For i = 0 To 0
        If i <> 0 Then
            grdMain(i).TextMatrix(0, 0) = " No " & i
        Else
            grdMain(i).TextMatrix(0, 0) = " No "
        End If
        grdMain(i).TextMatrix(0, 1) = "Description"
        grdMain(i).TextMatrix(0, 2) = "Total "
        If i = 0 Then
            grdMain(i).ColWidth(0) = grdMain(i).Width * 0.13
            grdMain(i).ColWidth(1) = grdMain(i).Width * 0.55
            grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
            grdMain(i).ColWidth(18) = grdMain(i).Width * 0.07
            grdMain(i).ColDataType(18) = flexDTBoolean
        Else
            grdMain(i).ColWidth(0) = grdMain(i).Width * 0.15
            grdMain(i).ColWidth(1) = grdMain(i).Width * 0.6
            grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
            grdMain(i).Rows = 1
        End If
        grdMain(i).ColAlignment(0) = flexAlignLeftCenter
        grdMain(i).ColAlignment(1) = flexAlignLeftCenter
        grdMain(i).ColAlignment(2) = flexAlignRightCenter
        grdMain(i).ColHidden(3) = True
        grdMain(i).ColHidden(4) = True
        grdMain(i).ColHidden(5) = True
        grdMain(i).ColHidden(6) = True
        grdMain(i).ColHidden(7) = True
        grdMain(i).ColHidden(8) = True
        grdMain(i).ColHidden(9) = True
        grdMain(i).ColHidden(10) = True
        grdMain(i).ColHidden(11) = True
        grdMain(i).ColHidden(12) = True
        grdMain(i).ColHidden(13) = True
        grdMain(i).ColHidden(14) = True
        grdMain(i).ColHidden(15) = True
        grdMain(i).ColHidden(16) = True
        grdMain(i).ColHidden(17) = True
        If i = 0 Then grdMain(i).ColHidden(18) = False
        grdMain(i).Cell(flexcpForeColor, 0, 0, 0, 2) = 0
    Next i
End Sub
Private Sub LoadTabs(Action)
    grdTab.Rows = 0
    cmdTable(0).Caption = ""
    cmdTable(0).Picture = ""
    DoEvents
    Select Case Action
        Case 0
            ActiveReadServer "Select * from Tab_Listing_View where Locked=0 and Tab_No <> " & TillData.TabNo
    End Select
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdTab.Rows = grdTab.Rows + 1
        If i < 17 And Not rs.EOF Then
            cmdTable(i).Caption = rs.Fields("Tab_Name") & ""
            cmdTable(i).Tag = rs.Fields("Tab_No")
            If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
            grdTab.Row = grdTab.Rows - 1
            grdTab.TextMatrix(grdTab.Rows - 1, 0) = rs.Fields("Tab_No")
            grdTab.TextMatrix(grdTab.Rows - 1, 1) = rs.Fields("Covers")
            grdTab.TextMatrix(grdTab.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTab.TextMatrix(grdTab.Rows - 1, 3) = rs.Fields("Locked")
            grdTab.TextMatrix(grdTab.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTab.TextMatrix(grdTab.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            grdTab.TextMatrix(grdTab.Rows - 1, 6) = rs.Fields("Tab_Name")
            If Action = 1 Or Action = 3 Then
                cmdTable(i).TextDescrCB.OffsetY = -10
                cmdTable(i).TextDescrCB.ColorNormal = &H800000
                cmdTable(i).TextDescrCB.Text = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                If rs.Fields("Locked") = True And rs.Fields("Workstation_No") <> Workstation_No Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTable(i).TextDescrCT.Text = ""
                End If
            Else
                cmdTable(i).TextDescrCB.Text = ""
                If rs.Fields("Previous_Owner") = rs.Fields("User_No") Then
                    cmdTable(i).TextDescrCB.Text = ""
                Else
                    cmdTable(i).TextDescrCB.OffsetY = -10
                    cmdTable(i).TextDescrCB.ColorNormal = &H800000
                    cmdTable(i).TextDescrCB.Text = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    If rs.Fields("Previous_Name") & "" <> "" Then
                        If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                            cmdTable(i).TextDescrCB.Text = "From > " & rs.Fields("Previous_Name") & ""
                        End If
                    End If
                End If
                If rs.Fields("Locked") = True And rs.Fields("Workstation_No") <> Workstation_No Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTable(i).TextDescrCT.Text = ""
                End If
            End If
        Else
            If b = 0 Then
                grdTab.TextMatrix(grdTab.Rows - 1, 0) = "ßArrow"
                grdTab.Rows = grdTab.Rows + 1
                If i = 17 Then
                    cmdTable(17).Caption = ""
                    cmdTable(17).Picture = App.Path & "\icons\downArr.bmp"
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    If cmdTable(17).Visible = False Then cmdTable(17).Visible = True
                End If
            End If
            b = b + 1
            grdTab.TextMatrix(grdTab.Rows - 1, 0) = rs.Fields("Tab_No")
            grdTab.TextMatrix(grdTab.Rows - 1, 1) = rs.Fields("Covers")
            grdTab.TextMatrix(grdTab.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTab.TextMatrix(grdTab.Rows - 1, 3) = rs.Fields("Locked")
            grdTab.TextMatrix(grdTab.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTab.TextMatrix(grdTab.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            grdTab.TextMatrix(grdTab.Rows - 1, 6) = rs.Fields("Tab_Name")
            If b = 16 Then b = 0
        End If
        rs.MoveNext
    Wend
    Select Case Action
        Case 0, 2
            If rs.RecordCount = 1 Then
                lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Tab"
            Else
                lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Tabs"
            End If
        Case 1, 3
            If rs.RecordCount = 1 Then
                lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Tab"
            Else
                lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Tabs"
            End If
    End Select
    rs.Close
    For b = i + 1 To cmdTable.Count - 1
       cmdTable(b).Caption = "0"
       cmdTable(b).Visible = False
    Next b
End Sub
Private Sub LoadTables(Action)
    grdTable.Rows = 0
    cmdTable(0).Caption = ""
    cmdTable(0).Picture = ""
    DoEvents
    Select Case Action
        Case 0
            ActiveReadServer "Select * from Table_Listing_View where Locked=0 and Table_No <> " & TillData.TableNo & " and Table_No <> 9999"
    End Select
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdTable.Rows = grdTable.Rows + 1
        If i < 17 And Not rs.EOF Then
            cmdTable(i).Caption = "Table No: " & rs.Fields("Table_No")
            cmdTable(i).Tag = rs.Fields("Covers")
            If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
            grdTable.Row = grdTable.Rows - 1
            grdTable.TextMatrix(grdTable.Rows - 1, 0) = rs.Fields("Table_No")
            grdTable.TextMatrix(grdTable.Rows - 1, 1) = rs.Fields("Covers")
            grdTable.TextMatrix(grdTable.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTable.TextMatrix(grdTable.Rows - 1, 3) = rs.Fields("Locked")
            grdTable.TextMatrix(grdTable.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTable.TextMatrix(grdTable.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            cmdTable(i).TextDescrCB.OffsetY = -10
            cmdTable(i).TextDescrCB.ColorNormal = &H800000
            cmdTable(i).TextDescrCB.Text = grdTable.TextMatrix(grdTable.Row, 2)
            If grdTable.TextMatrix(grdTable.Row, 3) = "True" And Workstation_No <> Val(grdTable.TextMatrix(grdTable.Row, 4)) Then
                cmdTable(i).TextDescrCT.OffsetY = 12
                cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                cmdTable(i).TextDescrCT.Text = "In Use"
            Else
                cmdTable(i).TextDescrCT.Text = ""
            End If
        Else
            If b = 0 Then
                grdTable.TextMatrix(grdTable.Rows - 1, 0) = "ßArrow"
                grdTable.Rows = grdTable.Rows + 1
                If i = 17 Then
                    cmdTable(17).Caption = ""
                    cmdTable(17).Picture = App.Path & "\icons\downArr.bmp"
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    If cmdTable(17).Visible = False Then cmdTable(17).Visible = True
                End If
            End If
            b = b + 1
            grdTable.TextMatrix(grdTable.Rows - 1, 0) = rs.Fields("Table_No")
            grdTable.TextMatrix(grdTable.Rows - 1, 1) = rs.Fields("Covers")
            grdTable.TextMatrix(grdTable.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTable.TextMatrix(grdTable.Rows - 1, 3) = rs.Fields("Locked")
            grdTable.TextMatrix(grdTable.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTable.TextMatrix(grdTable.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            If b = 16 Then b = 0
        End If
        rs.MoveNext
    Wend
    Select Case Action
        Case 0
            If rs.RecordCount = 1 Then
                lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Table"
            Else
                lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Tables"
            End If
        Case 1
            If rs.RecordCount = 1 Then
                lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Tab"
            Else
                lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Tabs"
            End If
    End Select
    rs.Close
    For b = i + 1 To cmdTable.Count - 1
       cmdTable(b).Caption = "0"
       cmdTable(b).Visible = False
    Next b
End Sub

Private Sub grdMain_Click(Index As Integer)
    If grdMain(Index).Tag = "" Then
        Select Case grdMain(0).ValueMatrix(grdMain(0).Row, 18)
            Case True
                grdMain(0).TextMatrix(grdMain(0).Row, 18) = 0
                grdMain(0).Cell(flexcpBackColor, grdMain(0).Row, 0, grdMain(0).Row, 18) = vbWhite
            Case False
                grdMain(0).TextMatrix(grdMain(0).Row, 18) = 1
                grdMain(0).Cell(flexcpBackColor, grdMain(0).Row, 0, grdMain(0).Row, 18) = &HC0FFC0
        End Select
    End If
    Subtotal = 0
    For i = 1 To grdMain(0).Rows - 1
        Subtotal = Subtotal + grdMain(Index).ValueMatrix(i, 2)
    Next i
    Select Case Panel_no
        Case 1
            lblTotal.Caption = "Subtotal: " & Format(Subtotal, "0.00")
        Case 2
            lblTotal.Caption = "Subtotal: " & Format(Subtotal, "0.00")
    End Select
End Sub
Private Sub grdMain_BeforeMouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    If y < grdMain(Index).Rows * grdMain(Index).RowHeightMin Then
        grdMain(Index).Tag = ""
    Else
        grdMain(Index).Tag = "1"
    End If
End Sub
Private Sub Update_Sale()
    Select Case Panel_no
        Case 1
            ActiveReadServer "Select * from Table_Listing_View where Table_No = " & TillData.TableNo
            If rs.RecordCount > 0 Then
                User_No = rs.Fields("User_No")
            Else
                User_No = UserRecord.User_Number
            End If
            rs.Close
            ActiveUpdateServer "Delete from Table_Listing where Table_No= " & TillData.TableNo
            DoEvents
            If grdMain(0).Rows > 1 Then
                For ib = 1 To grdMain(0).Rows - 1
                    ActiveUpdateServer "INSERT INTO [Table_Listing]([Table_No],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value, [Table_name])" & _
                    " VALUES('" & Val(TillData.TableNo) & "','" & TillData.Covers & "','" & User_No & "','" & Workstation_No & "','" & grdMain(0).TextMatrix(ib, 0) & "','" & grdMain(0).TextMatrix(ib, 1) & "','" & _
                    grdMain(0).TextMatrix(ib, 2) & "','" & grdMain(0).TextMatrix(ib, 3) & "','" & grdMain(0).TextMatrix(ib, 4) & "','" & grdMain(0).TextMatrix(ib, 5) & "','" & grdMain(0).TextMatrix(ib, 6) & "','" & grdMain(0).TextMatrix(ib, 8) & "','" & grdMain(0).TextMatrix(ib, 9) & "','" & grdMain(0).TextMatrix(ib, 10) & "','" & grdMain(0).TextMatrix(ib, 11) & "','" & grdMain(0).TextMatrix(ib, 12) & "','" & grdMain(0).TextMatrix(ib, 13) & "','" & grdMain(0).TextMatrix(ib, 14) & "','" & grdMain(0).TextMatrix(ib, 7) & "'," & TillData.DocNo & ",0,'" & grdMain(0).TextMatrix(ib, 17) & "',0," & grdMain(0).ValueMatrix(ib, 19) & ",'" & TillData.Table_Name & "')"
                Next ib
            End If
        Case 2
            ActiveReadServer "Select * from Tab_Listing_View where Tab_No = " & TillData.TabNo
            If rs.RecordCount > 0 Then
                User_No = rs.Fields("User_No")
            Else
                User_No = UserRecord.User_Number
            End If
            rs.Close
            ActiveUpdateServer "Delete from Tab_Listing where Tab_No= " & TillData.TabNo
            DoEvents
            If grdMain(0).Rows > 1 Then
                For ib = 1 To grdMain(0).Rows - 1
                    ActiveUpdateServer "INSERT INTO [Tab_Listing]([Tab_No],[Tab_Name],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value)" & _
                    " VALUES('" & Val(TillData.TabNo) & "','" & TillData.TabName & "','" & TillData.Covers & "','" & User_No & "','" & Workstation_No & "','" & grdMain(0).TextMatrix(ib, 0) & "','" & grdMain(0).TextMatrix(ib, 1) & "','" & _
                    grdMain(0).TextMatrix(ib, 2) & "','" & grdMain(0).TextMatrix(ib, 3) & "','" & grdMain(0).TextMatrix(ib, 4) & "','" & grdMain(0).TextMatrix(ib, 5) & "','" & grdMain(0).TextMatrix(ib, 6) & "','" & grdMain(0).TextMatrix(ib, 8) & "','" & grdMain(0).TextMatrix(ib, 9) & "','" & grdMain(0).TextMatrix(ib, 10) & "','" & grdMain(0).TextMatrix(ib, 11) & "','" & grdMain(0).TextMatrix(ib, 12) & "','" & grdMain(0).TextMatrix(ib, 13) & "','" & grdMain(0).TextMatrix(ib, 14) & "','" & grdMain(0).TextMatrix(ib, 7) & "'," & TillData.DocNo & ",0,'" & grdMain(0).TextMatrix(ib, 17) & "',0," & grdMain(0).ValueMatrix(ib, 19) & ")"
                Next ib
            End If
    End Select
End Sub
Private Sub scrolTimer_Timer()
    scrolTimer.Interval = 50
    Select Case scrolTimer.Tag
        Case "0"
            If grdMain(0).Row <> 1 Then
                grdMain(0).Row = grdMain(0).Row - 1
            End If
        Case "1"
            If grdMain(0).Row <> grdMain(0).Rows - 1 Then
                grdMain(0).Row = grdMain(0).Row + 1
            End If
    End Select
    grdMain(0).ShowCell grdMain(0).Row, 0

End Sub
