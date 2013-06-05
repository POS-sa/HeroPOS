VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form Startup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "HeroPOS First time startup"
   ClientHeight    =   9435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7860
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Startup.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Startup.frx":1CCA
   ScaleHeight     =   9435
   ScaleWidth      =   7860
   StartUpPosition =   1  'CenterOwner
   Begin BTNENHLib4.BtnEnh BtnEnh4 
      Height          =   885
      Left            =   2130
      TabIndex        =   6
      Top             =   8280
      Width           =   3765
      _Version        =   524298
      _ExtentX        =   6641
      _ExtentY        =   1561
      _StockProps     =   66
      Caption         =   "Save and Restart"
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
      BackColorContainer=   8438015
      SpecialEffect   =   3
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "Startup.frx":8DB0
      textLT          =   "Startup.frx":8E30
      textCT          =   "Startup.frx":8E48
      textRT          =   "Startup.frx":8E60
      textLM          =   "Startup.frx":8E78
      textRM          =   "Startup.frx":8E90
      textLB          =   "Startup.frx":8EA8
      textCB          =   "Startup.frx":8EC0
      textRB          =   "Startup.frx":8ED8
      colorBack       =   "Startup.frx":8EF0
      colorIntern     =   "Startup.frx":8F1A
      colorMO         =   "Startup.frx":8F44
      colorFocus      =   "Startup.frx":8F6E
      colorDisabled   =   "Startup.frx":8F98
      colorPressed    =   "Startup.frx":8FC2
   End
   Begin BTNENHLib4.BtnEnh BtnEnh2 
      Height          =   585
      Left            =   2130
      TabIndex        =   5
      Top             =   4560
      Width           =   3705
      _Version        =   524298
      _ExtentX        =   6535
      _ExtentY        =   1032
      _StockProps     =   66
      Caption         =   "Database Name"
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
      BackColorContainer=   8438015
      SpecialEffect   =   3
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "Startup.frx":8FEC
      textLT          =   "Startup.frx":9066
      textCT          =   "Startup.frx":907E
      textRT          =   "Startup.frx":9096
      textLM          =   "Startup.frx":90AE
      textRM          =   "Startup.frx":90C6
      textLB          =   "Startup.frx":90DE
      textCB          =   "Startup.frx":90F6
      textRB          =   "Startup.frx":910E
      colorBack       =   "Startup.frx":9126
      colorIntern     =   "Startup.frx":9150
      colorMO         =   "Startup.frx":917A
      colorFocus      =   "Startup.frx":91A4
      colorDisabled   =   "Startup.frx":91CE
      colorPressed    =   "Startup.frx":91F8
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   585
      Left            =   2130
      TabIndex        =   4
      Top             =   3270
      Width           =   3705
      _Version        =   524298
      _ExtentX        =   6535
      _ExtentY        =   1032
      _StockProps     =   66
      Caption         =   "Server Name"
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
      BackColorContainer=   8438015
      SpecialEffect   =   3
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "Startup.frx":9222
      textLT          =   "Startup.frx":9298
      textCT          =   "Startup.frx":92B0
      textRT          =   "Startup.frx":92C8
      textLM          =   "Startup.frx":92E0
      textRM          =   "Startup.frx":92F8
      textLB          =   "Startup.frx":9310
      textCB          =   "Startup.frx":9328
      textRB          =   "Startup.frx":9340
      colorBack       =   "Startup.frx":9358
      colorIntern     =   "Startup.frx":9382
      colorMO         =   "Startup.frx":93AC
      colorFocus      =   "Startup.frx":93D6
      colorDisabled   =   "Startup.frx":9400
      colorPressed    =   "Startup.frx":942A
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   885
      Left            =   330
      TabIndex        =   1
      Top             =   270
      Width           =   6315
      _Version        =   524298
      _ExtentX        =   11139
      _ExtentY        =   1561
      _StockProps     =   66
      Caption         =   "HeroPOS first time initialization..."
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
      BackColorContainer=   8438015
      SpecialEffect   =   3
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "Startup.frx":9454
      textLT          =   "Startup.frx":94FC
      textCT          =   "Startup.frx":9514
      textRT          =   "Startup.frx":952C
      textLM          =   "Startup.frx":9544
      textRM          =   "Startup.frx":955C
      textLB          =   "Startup.frx":9574
      textCB          =   "Startup.frx":958C
      textRB          =   "Startup.frx":95A4
      colorBack       =   "Startup.frx":95BC
      colorIntern     =   "Startup.frx":95E6
      colorMO         =   "Startup.frx":9610
      colorFocus      =   "Startup.frx":963A
      colorDisabled   =   "Startup.frx":9664
      colorPressed    =   "Startup.frx":968E
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   900
      Index           =   0
      Left            =   6810
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   270
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1588
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
   Begin BTNENHLib4.BtnEnh BtnEnh3 
      Height          =   585
      Left            =   2130
      TabIndex        =   8
      Top             =   5850
      Width           =   3705
      _Version        =   524298
      _ExtentX        =   6535
      _ExtentY        =   1032
      _StockProps     =   66
      Caption         =   "Password"
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
      BackColorContainer=   8438015
      SpecialEffect   =   3
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "Startup.frx":96B8
      textLT          =   "Startup.frx":9728
      textCT          =   "Startup.frx":9740
      textRT          =   "Startup.frx":9758
      textLM          =   "Startup.frx":9770
      textRM          =   "Startup.frx":9788
      textLB          =   "Startup.frx":97A0
      textCB          =   "Startup.frx":97B8
      textRB          =   "Startup.frx":97D0
      colorBack       =   "Startup.frx":97E8
      colorIntern     =   "Startup.frx":9812
      colorMO         =   "Startup.frx":983C
      colorFocus      =   "Startup.frx":9866
      colorDisabled   =   "Startup.frx":9890
      colorPressed    =   "Startup.frx":98BA
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   585
      Left            =   2130
      TabIndex        =   9
      Top             =   6510
      Width           =   3705
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6535;1032"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   765
      Left            =   690
      TabIndex        =   7
      Top             =   1440
      Width           =   6525
      ForeColor       =   0
      BackColor       =   12632256
      VariousPropertyBits=   8388627
      Caption         =   "This screen will only appear on first time installation of HeroPOS."
      Size            =   "11509;1349"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   585
      Left            =   2130
      TabIndex        =   3
      Top             =   5190
      Width           =   3705
      VariousPropertyBits=   746604571
      Size            =   "6535;1032"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   585
      Left            =   2130
      TabIndex        =   2
      Top             =   3900
      Width           =   3705
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "6535;1032"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "Startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnEnh4_Click()

            SaveSetting Trim(gblApp_Name), "Show_Splash", "Value", 1
            SaveSetting Trim(gblApp_Name), "Server", "Server", TextBox1.Text
            SaveSetting Trim(gblApp_Name), "Server", "SQL_User", "sa"
            SaveSetting Trim(gblApp_Name), "Server", "SQL_Password", TextBox3.Text
            SaveSetting Trim(gblApp_Name), "Server", "SQL_Database", TextBox2.Text
            SaveSetting Trim(gblApp_Name), "Logs", "Main_Log", App.Path & "\Logs"
            SaveSetting Trim(gblApp_Name), "Logs", "Error_Log", App.Path & "\Logs"
            
            Unload Me
            End
      
End Sub

Private Sub cmdCancel_Click(Index As Integer)
Unload Me
End
End Sub


Private Sub Form_Initialize()
Dim fso As New FileSystemObject
    If fso.FolderExists(App.Path & "\Logs\") = False Then fso.CreateFolder (App.Path & "\Logs")
    
End Sub
