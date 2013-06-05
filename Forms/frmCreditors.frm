VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmCreditors 
   BorderStyle     =   0  'None
   Caption         =   "Creditor Pay Out..."
   ClientHeight    =   9540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   Icon            =   "frmCreditors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCreditors.frx":000C
   ScaleHeight     =   9540
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   870
      Left            =   6750
      TabIndex        =   0
      Top             =   330
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   192
      Caption         =   "r"
      CaptionOffsetY  =   2
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VB.PictureBox picHold 
      Height          =   705
      Left            =   6930
      ScaleHeight     =   645
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   420
      Width           =   555
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   3
      Left            =   330
      TabIndex        =   1
      Top             =   2625
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F697E
      textLT          =   "frmCreditors.frx":F6996
      textCT          =   "frmCreditors.frx":F69AE
      textRT          =   "frmCreditors.frx":F69C6
      textLM          =   "frmCreditors.frx":F69DE
      textRM          =   "frmCreditors.frx":F69F6
      textLB          =   "frmCreditors.frx":F6A0E
      textCB          =   "frmCreditors.frx":F6A26
      textRB          =   "frmCreditors.frx":F6A3E
      colorBack       =   "frmCreditors.frx":F6A56
      colorIntern     =   "frmCreditors.frx":F6A80
      colorMO         =   "frmCreditors.frx":F6AAA
      colorFocus      =   "frmCreditors.frx":F6AD4
      colorDisabled   =   "frmCreditors.frx":F6AFE
      colorPressed    =   "frmCreditors.frx":F6B28
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   5
      Left            =   5160
      TabIndex        =   2
      Top             =   2625
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F6B52
      textLT          =   "frmCreditors.frx":F6B6A
      textCT          =   "frmCreditors.frx":F6B82
      textRT          =   "frmCreditors.frx":F6B9A
      textLM          =   "frmCreditors.frx":F6BB2
      textRM          =   "frmCreditors.frx":F6BCA
      textLB          =   "frmCreditors.frx":F6BE2
      textCB          =   "frmCreditors.frx":F6BFA
      textRB          =   "frmCreditors.frx":F6C12
      colorBack       =   "frmCreditors.frx":F6C2A
      colorIntern     =   "frmCreditors.frx":F6C54
      colorMO         =   "frmCreditors.frx":F6C7E
      colorFocus      =   "frmCreditors.frx":F6CA8
      colorDisabled   =   "frmCreditors.frx":F6CD2
      colorPressed    =   "frmCreditors.frx":F6CFC
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   6
      Left            =   330
      TabIndex        =   3
      Top             =   3930
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F6D26
      textLT          =   "frmCreditors.frx":F6D3E
      textCT          =   "frmCreditors.frx":F6D56
      textRT          =   "frmCreditors.frx":F6D6E
      textLM          =   "frmCreditors.frx":F6D86
      textRM          =   "frmCreditors.frx":F6D9E
      textLB          =   "frmCreditors.frx":F6DB6
      textCB          =   "frmCreditors.frx":F6DCE
      textRB          =   "frmCreditors.frx":F6DE6
      colorBack       =   "frmCreditors.frx":F6DFE
      colorIntern     =   "frmCreditors.frx":F6E28
      colorMO         =   "frmCreditors.frx":F6E52
      colorFocus      =   "frmCreditors.frx":F6E7C
      colorDisabled   =   "frmCreditors.frx":F6EA6
      colorPressed    =   "frmCreditors.frx":F6ED0
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   8
      Left            =   5160
      TabIndex        =   4
      Top             =   3930
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F6EFA
      textLT          =   "frmCreditors.frx":F6F12
      textCT          =   "frmCreditors.frx":F6F2A
      textRT          =   "frmCreditors.frx":F6F42
      textLM          =   "frmCreditors.frx":F6F5A
      textRM          =   "frmCreditors.frx":F6F72
      textLB          =   "frmCreditors.frx":F6F8A
      textCB          =   "frmCreditors.frx":F6FA2
      textRB          =   "frmCreditors.frx":F6FBA
      colorBack       =   "frmCreditors.frx":F6FD2
      colorIntern     =   "frmCreditors.frx":F6FFC
      colorMO         =   "frmCreditors.frx":F7026
      colorFocus      =   "frmCreditors.frx":F7050
      colorDisabled   =   "frmCreditors.frx":F707A
      colorPressed    =   "frmCreditors.frx":F70A4
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   9
      Left            =   330
      TabIndex        =   5
      Top             =   5235
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F70CE
      textLT          =   "frmCreditors.frx":F70E6
      textCT          =   "frmCreditors.frx":F70FE
      textRT          =   "frmCreditors.frx":F7116
      textLM          =   "frmCreditors.frx":F712E
      textRM          =   "frmCreditors.frx":F7146
      textLB          =   "frmCreditors.frx":F715E
      textCB          =   "frmCreditors.frx":F7176
      textRB          =   "frmCreditors.frx":F718E
      colorBack       =   "frmCreditors.frx":F71A6
      colorIntern     =   "frmCreditors.frx":F71D0
      colorMO         =   "frmCreditors.frx":F71FA
      colorFocus      =   "frmCreditors.frx":F7224
      colorDisabled   =   "frmCreditors.frx":F724E
      colorPressed    =   "frmCreditors.frx":F7278
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   11
      Left            =   5160
      TabIndex        =   6
      Top             =   5235
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F72A2
      textLT          =   "frmCreditors.frx":F72BA
      textCT          =   "frmCreditors.frx":F72D2
      textRT          =   "frmCreditors.frx":F72EA
      textLM          =   "frmCreditors.frx":F7302
      textRM          =   "frmCreditors.frx":F731A
      textLB          =   "frmCreditors.frx":F7332
      textCB          =   "frmCreditors.frx":F734A
      textRB          =   "frmCreditors.frx":F7362
      colorBack       =   "frmCreditors.frx":F737A
      colorIntern     =   "frmCreditors.frx":F73A4
      colorMO         =   "frmCreditors.frx":F73CE
      colorFocus      =   "frmCreditors.frx":F73F8
      colorDisabled   =   "frmCreditors.frx":F7422
      colorPressed    =   "frmCreditors.frx":F744C
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   12
      Left            =   330
      TabIndex        =   7
      Top             =   6540
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F7476
      textLT          =   "frmCreditors.frx":F748E
      textCT          =   "frmCreditors.frx":F74A6
      textRT          =   "frmCreditors.frx":F74BE
      textLM          =   "frmCreditors.frx":F74D6
      textRM          =   "frmCreditors.frx":F74EE
      textLB          =   "frmCreditors.frx":F7506
      textCB          =   "frmCreditors.frx":F751E
      textRB          =   "frmCreditors.frx":F7536
      colorBack       =   "frmCreditors.frx":F754E
      colorIntern     =   "frmCreditors.frx":F7578
      colorMO         =   "frmCreditors.frx":F75A2
      colorFocus      =   "frmCreditors.frx":F75CC
      colorDisabled   =   "frmCreditors.frx":F75F6
      colorPressed    =   "frmCreditors.frx":F7620
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   14
      Left            =   5160
      TabIndex        =   8
      Top             =   6540
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F764A
      textLT          =   "frmCreditors.frx":F7662
      textCT          =   "frmCreditors.frx":F767A
      textRT          =   "frmCreditors.frx":F7692
      textLM          =   "frmCreditors.frx":F76AA
      textRM          =   "frmCreditors.frx":F76C2
      textLB          =   "frmCreditors.frx":F76DA
      textCB          =   "frmCreditors.frx":F76F2
      textRB          =   "frmCreditors.frx":F770A
      colorBack       =   "frmCreditors.frx":F7722
      colorIntern     =   "frmCreditors.frx":F774C
      colorMO         =   "frmCreditors.frx":F7776
      colorFocus      =   "frmCreditors.frx":F77A0
      colorDisabled   =   "frmCreditors.frx":F77CA
      colorPressed    =   "frmCreditors.frx":F77F4
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1335
      Index           =   15
      Left            =   330
      TabIndex        =   9
      Top             =   7845
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2355
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F781E
      textLT          =   "frmCreditors.frx":F7836
      textCT          =   "frmCreditors.frx":F784E
      textRT          =   "frmCreditors.frx":F7866
      textLM          =   "frmCreditors.frx":F787E
      textRM          =   "frmCreditors.frx":F7896
      textLB          =   "frmCreditors.frx":F78AE
      textCB          =   "frmCreditors.frx":F78C6
      textRB          =   "frmCreditors.frx":F78DE
      colorBack       =   "frmCreditors.frx":F78F6
      colorIntern     =   "frmCreditors.frx":F7920
      colorMO         =   "frmCreditors.frx":F794A
      colorFocus      =   "frmCreditors.frx":F7974
      colorDisabled   =   "frmCreditors.frx":F799E
      colorPressed    =   "frmCreditors.frx":F79C8
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1335
      Index           =   17
      Left            =   5160
      TabIndex        =   10
      Top             =   7845
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2355
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F79F2
      textLT          =   "frmCreditors.frx":F7A0A
      textCT          =   "frmCreditors.frx":F7A22
      textRT          =   "frmCreditors.frx":F7A3A
      textLM          =   "frmCreditors.frx":F7A52
      textRM          =   "frmCreditors.frx":F7A6A
      textLB          =   "frmCreditors.frx":F7A82
      textCB          =   "frmCreditors.frx":F7A9A
      textRB          =   "frmCreditors.frx":F7AB2
      colorBack       =   "frmCreditors.frx":F7ACA
      colorIntern     =   "frmCreditors.frx":F7AF4
      colorMO         =   "frmCreditors.frx":F7B1E
      colorFocus      =   "frmCreditors.frx":F7B48
      colorDisabled   =   "frmCreditors.frx":F7B72
      colorPressed    =   "frmCreditors.frx":F7B9C
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   2
      Left            =   5160
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F7BC6
      textLT          =   "frmCreditors.frx":F7BDE
      textCT          =   "frmCreditors.frx":F7BF6
      textRT          =   "frmCreditors.frx":F7C0E
      textLM          =   "frmCreditors.frx":F7C26
      textRM          =   "frmCreditors.frx":F7C3E
      textLB          =   "frmCreditors.frx":F7C56
      textCB          =   "frmCreditors.frx":F7C6E
      textRB          =   "frmCreditors.frx":F7C86
      colorBack       =   "frmCreditors.frx":F7C9E
      colorIntern     =   "frmCreditors.frx":F7CC8
      colorMO         =   "frmCreditors.frx":F7CF2
      colorFocus      =   "frmCreditors.frx":F7D1C
      colorDisabled   =   "frmCreditors.frx":F7D46
      colorPressed    =   "frmCreditors.frx":F7D70
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   1
      Left            =   2745
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F7D9A
      textLT          =   "frmCreditors.frx":F7DB2
      textCT          =   "frmCreditors.frx":F7DCA
      textRT          =   "frmCreditors.frx":F7DE2
      textLM          =   "frmCreditors.frx":F7DFA
      textRM          =   "frmCreditors.frx":F7E12
      textLB          =   "frmCreditors.frx":F7E2A
      textCB          =   "frmCreditors.frx":F7E42
      textRB          =   "frmCreditors.frx":F7E5A
      colorBack       =   "frmCreditors.frx":F7E72
      colorIntern     =   "frmCreditors.frx":F7E9C
      colorMO         =   "frmCreditors.frx":F7EC6
      colorFocus      =   "frmCreditors.frx":F7EF0
      colorDisabled   =   "frmCreditors.frx":F7F1A
      colorPressed    =   "frmCreditors.frx":F7F44
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   4
      Left            =   2745
      TabIndex        =   13
      Top             =   2625
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F7F6E
      textLT          =   "frmCreditors.frx":F7F86
      textCT          =   "frmCreditors.frx":F7F9E
      textRT          =   "frmCreditors.frx":F7FB6
      textLM          =   "frmCreditors.frx":F7FCE
      textRM          =   "frmCreditors.frx":F7FE6
      textLB          =   "frmCreditors.frx":F7FFE
      textCB          =   "frmCreditors.frx":F8016
      textRB          =   "frmCreditors.frx":F802E
      colorBack       =   "frmCreditors.frx":F8046
      colorIntern     =   "frmCreditors.frx":F8070
      colorMO         =   "frmCreditors.frx":F809A
      colorFocus      =   "frmCreditors.frx":F80C4
      colorDisabled   =   "frmCreditors.frx":F80EE
      colorPressed    =   "frmCreditors.frx":F8118
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   7
      Left            =   2745
      TabIndex        =   14
      Top             =   3930
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F8142
      textLT          =   "frmCreditors.frx":F815A
      textCT          =   "frmCreditors.frx":F8172
      textRT          =   "frmCreditors.frx":F818A
      textLM          =   "frmCreditors.frx":F81A2
      textRM          =   "frmCreditors.frx":F81BA
      textLB          =   "frmCreditors.frx":F81D2
      textCB          =   "frmCreditors.frx":F81EA
      textRB          =   "frmCreditors.frx":F8202
      colorBack       =   "frmCreditors.frx":F821A
      colorIntern     =   "frmCreditors.frx":F8244
      colorMO         =   "frmCreditors.frx":F826E
      colorFocus      =   "frmCreditors.frx":F8298
      colorDisabled   =   "frmCreditors.frx":F82C2
      colorPressed    =   "frmCreditors.frx":F82EC
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   10
      Left            =   2745
      TabIndex        =   15
      Top             =   5235
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F8316
      textLT          =   "frmCreditors.frx":F832E
      textCT          =   "frmCreditors.frx":F8346
      textRT          =   "frmCreditors.frx":F835E
      textLM          =   "frmCreditors.frx":F8376
      textRM          =   "frmCreditors.frx":F838E
      textLB          =   "frmCreditors.frx":F83A6
      textCB          =   "frmCreditors.frx":F83BE
      textRB          =   "frmCreditors.frx":F83D6
      colorBack       =   "frmCreditors.frx":F83EE
      colorIntern     =   "frmCreditors.frx":F8418
      colorMO         =   "frmCreditors.frx":F8442
      colorFocus      =   "frmCreditors.frx":F846C
      colorDisabled   =   "frmCreditors.frx":F8496
      colorPressed    =   "frmCreditors.frx":F84C0
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   13
      Left            =   2745
      TabIndex        =   16
      Top             =   6540
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F84EA
      textLT          =   "frmCreditors.frx":F8502
      textCT          =   "frmCreditors.frx":F851A
      textRT          =   "frmCreditors.frx":F8532
      textLM          =   "frmCreditors.frx":F854A
      textRM          =   "frmCreditors.frx":F8562
      textLB          =   "frmCreditors.frx":F857A
      textCB          =   "frmCreditors.frx":F8592
      textRB          =   "frmCreditors.frx":F85AA
      colorBack       =   "frmCreditors.frx":F85C2
      colorIntern     =   "frmCreditors.frx":F85EC
      colorMO         =   "frmCreditors.frx":F8616
      colorFocus      =   "frmCreditors.frx":F8640
      colorDisabled   =   "frmCreditors.frx":F866A
      colorPressed    =   "frmCreditors.frx":F8694
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1335
      Index           =   16
      Left            =   2745
      TabIndex        =   17
      Top             =   7845
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2355
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F86BE
      textLT          =   "frmCreditors.frx":F86D6
      textCT          =   "frmCreditors.frx":F86EE
      textRT          =   "frmCreditors.frx":F8706
      textLM          =   "frmCreditors.frx":F871E
      textRM          =   "frmCreditors.frx":F8736
      textLB          =   "frmCreditors.frx":F874E
      textCB          =   "frmCreditors.frx":F8766
      textRB          =   "frmCreditors.frx":F877E
      colorBack       =   "frmCreditors.frx":F8796
      colorIntern     =   "frmCreditors.frx":F87C0
      colorMO         =   "frmCreditors.frx":F87EA
      colorFocus      =   "frmCreditors.frx":F8814
      colorDisabled   =   "frmCreditors.frx":F883E
      colorPressed    =   "frmCreditors.frx":F8868
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   0
      Left            =   330
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCreditors.frx":F8892
      textLT          =   "frmCreditors.frx":F88AA
      textCT          =   "frmCreditors.frx":F88C2
      textRT          =   "frmCreditors.frx":F88DA
      textLM          =   "frmCreditors.frx":F88F2
      textRM          =   "frmCreditors.frx":F890A
      textLB          =   "frmCreditors.frx":F8922
      textCB          =   "frmCreditors.frx":F893A
      textRB          =   "frmCreditors.frx":F8952
      colorBack       =   "frmCreditors.frx":F896A
      colorIntern     =   "frmCreditors.frx":F8994
      colorMO         =   "frmCreditors.frx":F89BE
      colorFocus      =   "frmCreditors.frx":F89E8
      colorDisabled   =   "frmCreditors.frx":F8A12
      colorPressed    =   "frmCreditors.frx":F8A3C
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdCreditors 
      Height          =   8370
      Left            =   8310
      TabIndex        =   20
      Top             =   180
      Visible         =   0   'False
      Width           =   2865
      _cx             =   5054
      _cy             =   14764
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
   Begin MSForms.Label lblTransfer 
      Height          =   765
      Left            =   510
      TabIndex        =   21
      Top             =   480
      Width           =   5085
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Select a Creditor to Pay"
      Size            =   "8969;1349"
      FontName        =   "Arial Narrow"
      FontHeight      =   480
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmCreditors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
