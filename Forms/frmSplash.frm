VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin BTNENHLib4.BtnEnh BtnEnh 
      Height          =   1995
      Left            =   0
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   15285
      _Version        =   524298
      _ExtentX        =   26961
      _ExtentY        =   3519
      _StockProps     =   66
      Caption         =   "BETA - TESTING"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   72
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
         Size            =   20.25
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
      BackColorContainer=   11131362
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":383C5
      textLT          =   "frmSplash.frx":38441
      textCT          =   "frmSplash.frx":38459
      textRT          =   "frmSplash.frx":38471
      textLM          =   "frmSplash.frx":38489
      textRM          =   "frmSplash.frx":384A1
      textLB          =   "frmSplash.frx":384B9
      textCB          =   "frmSplash.frx":384D1
      textRB          =   "frmSplash.frx":384E9
      colorBack       =   "frmSplash.frx":38501
      colorIntern     =   "frmSplash.frx":3852B
      colorMO         =   "frmSplash.frx":38555
      colorFocus      =   "frmSplash.frx":3857F
      colorDisabled   =   "frmSplash.frx":385A9
      colorPressed    =   "frmSplash.frx":385D3
      Style           =   3
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   2220
      Top             =   2430
   End
   Begin VB.Timer Timer5 
      Interval        =   60000
      Left            =   450
      Top             =   2850
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   4065
      Left            =   540
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   4005
      _Version        =   524298
      _ExtentX        =   7064
      _ExtentY        =   7170
      _StockProps     =   66
      Caption         =   "Please Sign on to Access "
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
      Shape           =   1
      CornerFactor    =   14
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   85
      LogPixels       =   96
      SpecialEffectFactor=   3
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":385FD
      textLT          =   "frmSplash.frx":3868F
      textCT          =   "frmSplash.frx":386A7
      textRT          =   "frmSplash.frx":386BF
      textLM          =   "frmSplash.frx":386D7
      textRM          =   "frmSplash.frx":386EF
      textLB          =   "frmSplash.frx":38707
      textCB          =   "frmSplash.frx":3871F
      textRB          =   "frmSplash.frx":38737
      colorBack       =   "frmSplash.frx":3874F
      colorIntern     =   "frmSplash.frx":38779
      colorMO         =   "frmSplash.frx":387A3
      colorFocus      =   "frmSplash.frx":387CD
      colorDisabled   =   "frmSplash.frx":387F7
      colorPressed    =   "frmSplash.frx":38821
      Orientation     =   2
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1065
      Index           =   18
      Left            =   540
      TabIndex        =   36
      Top             =   7650
      Width           =   4020
      _Version        =   524298
      _ExtentX        =   7082
      _ExtentY        =   1879
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":3884B
      textLT          =   "frmSplash.frx":38863
      textCT          =   "frmSplash.frx":3887B
      textRT          =   "frmSplash.frx":38893
      textLM          =   "frmSplash.frx":388AB
      textRM          =   "frmSplash.frx":388C3
      textLB          =   "frmSplash.frx":388DB
      textCB          =   "frmSplash.frx":388F3
      textRB          =   "frmSplash.frx":3890B
      colorBack       =   "frmSplash.frx":38923
      colorIntern     =   "frmSplash.frx":3894D
      colorMO         =   "frmSplash.frx":38977
      colorFocus      =   "frmSplash.frx":389A1
      colorDisabled   =   "frmSplash.frx":389CB
      colorPressed    =   "frmSplash.frx":389F5
      Orientation     =   3
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh BtnEnh3 
      Height          =   1575
      Left            =   390
      TabIndex        =   28
      Top             =   2820
      Width           =   4305
      _Version        =   524298
      _ExtentX        =   7594
      _ExtentY        =   2778
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      CornerFactor    =   50
      BackColorContainer=   11131362
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":38A1F
      textLT          =   "frmSplash.frx":38A37
      textCT          =   "frmSplash.frx":38A4F
      textRT          =   "frmSplash.frx":38A67
      textLM          =   "frmSplash.frx":38A7F
      textRM          =   "frmSplash.frx":38A97
      textLB          =   "frmSplash.frx":38AAF
      textCB          =   "frmSplash.frx":38AC7
      textRB          =   "frmSplash.frx":38ADF
      colorBack       =   "frmSplash.frx":38AF7
      colorIntern     =   "frmSplash.frx":38B21
      colorMO         =   "frmSplash.frx":38B4B
      colorFocus      =   "frmSplash.frx":38B75
      colorDisabled   =   "frmSplash.frx":38B9F
      colorPressed    =   "frmSplash.frx":38BC9
      Style           =   3
      Begin BTNENHLib4.BtnEnh BtnEnh4 
         Height          =   615
         Left            =   210
         TabIndex        =   29
         Top             =   810
         Width           =   3885
         _Version        =   524298
         _ExtentX        =   6853
         _ExtentY        =   1085
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         CornerFactor    =   80
         BackColorContainer=   11131362
         LogPixels       =   96
         Clickable       =   0   'False
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSplash.frx":38BF3
         textLT          =   "frmSplash.frx":38C0B
         textCT          =   "frmSplash.frx":38C23
         textRT          =   "frmSplash.frx":38C3B
         textLM          =   "frmSplash.frx":38C53
         textRM          =   "frmSplash.frx":38C6B
         textLB          =   "frmSplash.frx":38C83
         textCB          =   "frmSplash.frx":38C9B
         textRB          =   "frmSplash.frx":38CB3
         colorBack       =   "frmSplash.frx":38CCB
         colorIntern     =   "frmSplash.frx":38CF5
         colorMO         =   "frmSplash.frx":38D1F
         colorFocus      =   "frmSplash.frx":38D49
         colorDisabled   =   "frmSplash.frx":38D73
         colorPressed    =   "frmSplash.frx":38D9D
         Style           =   3
         Orientation     =   4
         LightDirection  =   5
         Begin VB.TextBox txtPassword 
            Alignment       =   2  'Center
            BackColor       =   &H00A9D9E2&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   30
            Top             =   90
            Width           =   3345
         End
      End
      Begin BTNENHLib4.BtnEnh fmUser 
         Height          =   675
         Left            =   210
         TabIndex        =   31
         Top             =   180
         Width           =   3885
         _Version        =   524298
         _ExtentX        =   6853
         _ExtentY        =   1191
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         CornerFactor    =   80
         BackColorContainer=   11131362
         LogPixels       =   96
         Clickable       =   0   'False
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSplash.frx":38DC7
         textLT          =   "frmSplash.frx":38DDF
         textCT          =   "frmSplash.frx":38DF7
         textRT          =   "frmSplash.frx":38E0F
         textLM          =   "frmSplash.frx":38E27
         textRM          =   "frmSplash.frx":38E3F
         textLB          =   "frmSplash.frx":38E57
         textCB          =   "frmSplash.frx":38E6F
         textRB          =   "frmSplash.frx":38E87
         colorBack       =   "frmSplash.frx":38E9F
         colorIntern     =   "frmSplash.frx":38EC9
         colorMO         =   "frmSplash.frx":38EF3
         colorFocus      =   "frmSplash.frx":38F1D
         colorDisabled   =   "frmSplash.frx":38F47
         colorPressed    =   "frmSplash.frx":38F71
         Style           =   3
         Orientation     =   2
         LightDirection  =   5
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8820
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   1740
      Top             =   0
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1005
      Index           =   15
      Left            =   540
      TabIndex        =   19
      Top             =   6660
      Width           =   4020
      _Version        =   524298
      _ExtentX        =   7091
      _ExtentY        =   1773
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":38F9B
      textLT          =   "frmSplash.frx":38FB3
      textCT          =   "frmSplash.frx":38FCB
      textRT          =   "frmSplash.frx":38FE3
      textLM          =   "frmSplash.frx":38FFB
      textRM          =   "frmSplash.frx":39013
      textLB          =   "frmSplash.frx":3902B
      textCB          =   "frmSplash.frx":39043
      textRB          =   "frmSplash.frx":3905B
      colorBack       =   "frmSplash.frx":39073
      colorIntern     =   "frmSplash.frx":3909D
      colorMO         =   "frmSplash.frx":390C7
      colorFocus      =   "frmSplash.frx":390F1
      colorDisabled   =   "frmSplash.frx":3911B
      colorPressed    =   "frmSplash.frx":39145
      Orientation     =   3
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   9
      Left            =   5010
      TabIndex        =   13
      Top             =   7800
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "CL"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":3916F
      textLT          =   "frmSplash.frx":391D3
      textCT          =   "frmSplash.frx":391EB
      textRT          =   "frmSplash.frx":39203
      textLM          =   "frmSplash.frx":3921B
      textRM          =   "frmSplash.frx":39233
      textLB          =   "frmSplash.frx":3924B
      textCB          =   "frmSplash.frx":39263
      textRB          =   "frmSplash.frx":3927B
      colorBack       =   "frmSplash.frx":39293
      colorIntern     =   "frmSplash.frx":392BD
      colorMO         =   "frmSplash.frx":392E7
      colorFocus      =   "frmSplash.frx":39311
      colorDisabled   =   "frmSplash.frx":3933B
      colorPressed    =   "frmSplash.frx":39365
      Orientation     =   8
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   10
      Left            =   6555
      TabIndex        =   14
      Top             =   7800
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "0"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":3938F
      textLT          =   "frmSplash.frx":393F1
      textCT          =   "frmSplash.frx":39409
      textRT          =   "frmSplash.frx":39421
      textLM          =   "frmSplash.frx":39439
      textRM          =   "frmSplash.frx":39451
      textLB          =   "frmSplash.frx":39469
      textCB          =   "frmSplash.frx":39481
      textRB          =   "frmSplash.frx":39499
      colorBack       =   "frmSplash.frx":394B1
      colorIntern     =   "frmSplash.frx":394DB
      colorMO         =   "frmSplash.frx":39505
      colorFocus      =   "frmSplash.frx":3952F
      colorDisabled   =   "frmSplash.frx":39559
      colorPressed    =   "frmSplash.frx":39583
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   11
      Left            =   8100
      TabIndex        =   15
      Top             =   7800
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "Exit"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Shape           =   1
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":395AD
      textLT          =   "frmSplash.frx":39615
      textCT          =   "frmSplash.frx":3962D
      textRT          =   "frmSplash.frx":39645
      textLM          =   "frmSplash.frx":3965D
      textRM          =   "frmSplash.frx":39675
      textLB          =   "frmSplash.frx":3968D
      textCB          =   "frmSplash.frx":396A5
      textRB          =   "frmSplash.frx":396BD
      colorBack       =   "frmSplash.frx":396D5
      colorIntern     =   "frmSplash.frx":396FF
      colorMO         =   "frmSplash.frx":39729
      colorFocus      =   "frmSplash.frx":39753
      colorDisabled   =   "frmSplash.frx":3977D
      colorPressed    =   "frmSplash.frx":397A7
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   6
      Left            =   5010
      TabIndex        =   10
      Top             =   6360
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "7"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":397D1
      textLT          =   "frmSplash.frx":39833
      textCT          =   "frmSplash.frx":3984B
      textRT          =   "frmSplash.frx":39863
      textLM          =   "frmSplash.frx":3987B
      textRM          =   "frmSplash.frx":39893
      textLB          =   "frmSplash.frx":398AB
      textCB          =   "frmSplash.frx":398C3
      textRB          =   "frmSplash.frx":398DB
      colorBack       =   "frmSplash.frx":398F3
      colorIntern     =   "frmSplash.frx":3991D
      colorMO         =   "frmSplash.frx":39947
      colorFocus      =   "frmSplash.frx":39971
      colorDisabled   =   "frmSplash.frx":3999B
      colorPressed    =   "frmSplash.frx":399C5
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   7
      Left            =   6555
      TabIndex        =   11
      Top             =   6360
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "8"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":399EF
      textLT          =   "frmSplash.frx":39A51
      textCT          =   "frmSplash.frx":39A69
      textRT          =   "frmSplash.frx":39A81
      textLM          =   "frmSplash.frx":39A99
      textRM          =   "frmSplash.frx":39AB1
      textLB          =   "frmSplash.frx":39AC9
      textCB          =   "frmSplash.frx":39AE1
      textRB          =   "frmSplash.frx":39AF9
      colorBack       =   "frmSplash.frx":39B11
      colorIntern     =   "frmSplash.frx":39B3B
      colorMO         =   "frmSplash.frx":39B65
      colorFocus      =   "frmSplash.frx":39B8F
      colorDisabled   =   "frmSplash.frx":39BB9
      colorPressed    =   "frmSplash.frx":39BE3
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   8
      Left            =   8100
      TabIndex        =   12
      Top             =   6360
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "9"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":39C0D
      textLT          =   "frmSplash.frx":39C6F
      textCT          =   "frmSplash.frx":39C87
      textRT          =   "frmSplash.frx":39C9F
      textLM          =   "frmSplash.frx":39CB7
      textRM          =   "frmSplash.frx":39CCF
      textLB          =   "frmSplash.frx":39CE7
      textCB          =   "frmSplash.frx":39CFF
      textRB          =   "frmSplash.frx":39D17
      colorBack       =   "frmSplash.frx":39D2F
      colorIntern     =   "frmSplash.frx":39D59
      colorMO         =   "frmSplash.frx":39D83
      colorFocus      =   "frmSplash.frx":39DAD
      colorDisabled   =   "frmSplash.frx":39DD7
      colorPressed    =   "frmSplash.frx":39E01
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   0
      Left            =   5010
      TabIndex        =   4
      Top             =   3480
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Shape           =   1
      Surface         =   1
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":39E2B
      textLT          =   "frmSplash.frx":39E8D
      textCT          =   "frmSplash.frx":39EA5
      textRT          =   "frmSplash.frx":39EBD
      textLM          =   "frmSplash.frx":39ED5
      textRM          =   "frmSplash.frx":39EED
      textLB          =   "frmSplash.frx":39F05
      textCB          =   "frmSplash.frx":39F1D
      textRB          =   "frmSplash.frx":39F35
      colorBack       =   "frmSplash.frx":39F4D
      colorIntern     =   "frmSplash.frx":39F77
      colorMO         =   "frmSplash.frx":39FA1
      colorFocus      =   "frmSplash.frx":39FCB
      colorDisabled   =   "frmSplash.frx":39FF5
      colorPressed    =   "frmSplash.frx":3A01F
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   1
      Left            =   6555
      TabIndex        =   5
      Top             =   3480
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "2"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":3A049
      textLT          =   "frmSplash.frx":3A0AB
      textCT          =   "frmSplash.frx":3A0C3
      textRT          =   "frmSplash.frx":3A0DB
      textLM          =   "frmSplash.frx":3A0F3
      textRM          =   "frmSplash.frx":3A10B
      textLB          =   "frmSplash.frx":3A123
      textCB          =   "frmSplash.frx":3A13B
      textRB          =   "frmSplash.frx":3A153
      colorBack       =   "frmSplash.frx":3A16B
      colorIntern     =   "frmSplash.frx":3A195
      colorMO         =   "frmSplash.frx":3A1BF
      colorFocus      =   "frmSplash.frx":3A1E9
      colorDisabled   =   "frmSplash.frx":3A213
      colorPressed    =   "frmSplash.frx":3A23D
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   2
      Left            =   8100
      TabIndex        =   6
      Top             =   3480
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "3"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Shape           =   1
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":3A267
      textLT          =   "frmSplash.frx":3A2C9
      textCT          =   "frmSplash.frx":3A2E1
      textRT          =   "frmSplash.frx":3A2F9
      textLM          =   "frmSplash.frx":3A311
      textRM          =   "frmSplash.frx":3A329
      textLB          =   "frmSplash.frx":3A341
      textCB          =   "frmSplash.frx":3A359
      textRB          =   "frmSplash.frx":3A371
      colorBack       =   "frmSplash.frx":3A389
      colorIntern     =   "frmSplash.frx":3A3B3
      colorMO         =   "frmSplash.frx":3A3DD
      colorFocus      =   "frmSplash.frx":3A407
      colorDisabled   =   "frmSplash.frx":3A431
      colorPressed    =   "frmSplash.frx":3A45B
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   3
      Left            =   5010
      TabIndex        =   7
      Top             =   4920
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "4"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":3A485
      textLT          =   "frmSplash.frx":3A4E7
      textCT          =   "frmSplash.frx":3A4FF
      textRT          =   "frmSplash.frx":3A517
      textLM          =   "frmSplash.frx":3A52F
      textRM          =   "frmSplash.frx":3A547
      textLB          =   "frmSplash.frx":3A55F
      textCB          =   "frmSplash.frx":3A577
      textRB          =   "frmSplash.frx":3A58F
      colorBack       =   "frmSplash.frx":3A5A7
      colorIntern     =   "frmSplash.frx":3A5D1
      colorMO         =   "frmSplash.frx":3A5FB
      colorFocus      =   "frmSplash.frx":3A625
      colorDisabled   =   "frmSplash.frx":3A64F
      colorPressed    =   "frmSplash.frx":3A679
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   4
      Left            =   6555
      TabIndex        =   8
      Top             =   4920
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "5"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":3A6A3
      textLT          =   "frmSplash.frx":3A705
      textCT          =   "frmSplash.frx":3A71D
      textRT          =   "frmSplash.frx":3A735
      textLM          =   "frmSplash.frx":3A74D
      textRM          =   "frmSplash.frx":3A765
      textLB          =   "frmSplash.frx":3A77D
      textCB          =   "frmSplash.frx":3A795
      textRB          =   "frmSplash.frx":3A7AD
      colorBack       =   "frmSplash.frx":3A7C5
      colorIntern     =   "frmSplash.frx":3A7EF
      colorMO         =   "frmSplash.frx":3A819
      colorFocus      =   "frmSplash.frx":3A843
      colorDisabled   =   "frmSplash.frx":3A86D
      colorPressed    =   "frmSplash.frx":3A897
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1455
      Index           =   5
      Left            =   8100
      TabIndex        =   9
      Top             =   4920
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2566
      _StockProps     =   66
      Caption         =   "6"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":3A8C1
      textLT          =   "frmSplash.frx":3A923
      textCT          =   "frmSplash.frx":3A93B
      textRT          =   "frmSplash.frx":3A953
      textLM          =   "frmSplash.frx":3A96B
      textRM          =   "frmSplash.frx":3A983
      textLB          =   "frmSplash.frx":3A99B
      textCB          =   "frmSplash.frx":3A9B3
      textRB          =   "frmSplash.frx":3A9CB
      colorBack       =   "frmSplash.frx":3A9E3
      colorIntern     =   "frmSplash.frx":3AA0D
      colorMO         =   "frmSplash.frx":3AA37
      colorFocus      =   "frmSplash.frx":3AA61
      colorDisabled   =   "frmSplash.frx":3AA8B
      colorPressed    =   "frmSplash.frx":3AAB5
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMess 
      Height          =   7035
      Left            =   10290
      TabIndex        =   3
      Top             =   3000
      Width           =   4455
      _cx             =   7858
      _cy             =   12409
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   4210752
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaper       =   "frmSplash.frx":3AADF
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5190
      Top             =   2940
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1005
      Index           =   12
      Left            =   540
      TabIndex        =   16
      Top             =   4680
      Width           =   1995
      _Version        =   524298
      _ExtentX        =   3519
      _ExtentY        =   1773
      _StockProps     =   66
      Caption         =   "Clock In"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A0609
      textLT          =   "frmSplash.frx":A0679
      textCT          =   "frmSplash.frx":A0691
      textRT          =   "frmSplash.frx":A06A9
      textLM          =   "frmSplash.frx":A06C1
      textRM          =   "frmSplash.frx":A06D9
      textLB          =   "frmSplash.frx":A06F1
      textCB          =   "frmSplash.frx":A0709
      textRB          =   "frmSplash.frx":A0721
      colorBack       =   "frmSplash.frx":A0739
      colorIntern     =   "frmSplash.frx":A0763
      colorMO         =   "frmSplash.frx":A078D
      colorFocus      =   "frmSplash.frx":A07B7
      colorDisabled   =   "frmSplash.frx":A07E1
      colorPressed    =   "frmSplash.frx":A080B
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   2
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1005
      Index           =   13
      Left            =   2520
      TabIndex        =   17
      Top             =   4680
      Width           =   2025
      _Version        =   524298
      _ExtentX        =   3572
      _ExtentY        =   1773
      _StockProps     =   66
      Caption         =   "Clock Out"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      CaptionWordWrapPerc=   97
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A0835
      textLT          =   "frmSplash.frx":A08A7
      textCT          =   "frmSplash.frx":A08BF
      textRT          =   "frmSplash.frx":A08D7
      textLM          =   "frmSplash.frx":A08EF
      textRM          =   "frmSplash.frx":A0907
      textLB          =   "frmSplash.frx":A091F
      textCB          =   "frmSplash.frx":A0937
      textRB          =   "frmSplash.frx":A094F
      colorBack       =   "frmSplash.frx":A0967
      colorIntern     =   "frmSplash.frx":A0991
      colorMO         =   "frmSplash.frx":A09BB
      colorFocus      =   "frmSplash.frx":A09E5
      colorDisabled   =   "frmSplash.frx":A0A0F
      colorPressed    =   "frmSplash.frx":A0A39
      Orientation     =   6
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1125
      Index           =   14
      Left            =   525
      TabIndex        =   18
      Top             =   8700
      Width           =   4035
      _Version        =   524298
      _ExtentX        =   7108
      _ExtentY        =   1984
      _StockProps     =   66
      Caption         =   "Ok"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A0A63
      textLT          =   "frmSplash.frx":A0AC7
      textCT          =   "frmSplash.frx":A0ADF
      textRT          =   "frmSplash.frx":A0AF7
      textLM          =   "frmSplash.frx":A0B0F
      textRM          =   "frmSplash.frx":A0B27
      textLB          =   "frmSplash.frx":A0B3F
      textCB          =   "frmSplash.frx":A0B57
      textRB          =   "frmSplash.frx":A0B6F
      colorBack       =   "frmSplash.frx":A0B87
      colorIntern     =   "frmSplash.frx":A0BB1
      colorMO         =   "frmSplash.frx":A0BDB
      colorFocus      =   "frmSplash.frx":A0C05
      colorDisabled   =   "frmSplash.frx":A0C2F
      colorPressed    =   "frmSplash.frx":A0C59
      Orientation     =   4
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1005
      Index           =   16
      Left            =   540
      TabIndex        =   20
      Top             =   5670
      Width           =   1995
      _Version        =   524298
      _ExtentX        =   3519
      _ExtentY        =   1773
      _StockProps     =   66
      Caption         =   "Cashup's"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A0C83
      textLT          =   "frmSplash.frx":A0CF3
      textCT          =   "frmSplash.frx":A0D0B
      textRT          =   "frmSplash.frx":A0D23
      textLM          =   "frmSplash.frx":A0D3B
      textRM          =   "frmSplash.frx":A0D53
      textLB          =   "frmSplash.frx":A0D6B
      textCB          =   "frmSplash.frx":A0D83
      textRB          =   "frmSplash.frx":A0D9B
      colorBack       =   "frmSplash.frx":A0DB3
      colorIntern     =   "frmSplash.frx":A0DDD
      colorMO         =   "frmSplash.frx":A0E07
      colorFocus      =   "frmSplash.frx":A0E31
      colorDisabled   =   "frmSplash.frx":A0E5B
      colorPressed    =   "frmSplash.frx":A0E85
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1005
      Index           =   17
      Left            =   2520
      TabIndex        =   24
      Top             =   5670
      Width           =   2025
      _Version        =   524298
      _ExtentX        =   3572
      _ExtentY        =   1773
      _StockProps     =   66
      Caption         =   "Stock Takes"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      BackColorContainer=   11131362
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A0EAF
      textLT          =   "frmSplash.frx":A0F25
      textCT          =   "frmSplash.frx":A0F3D
      textRT          =   "frmSplash.frx":A0F55
      textLM          =   "frmSplash.frx":A0F6D
      textRM          =   "frmSplash.frx":A0F85
      textLB          =   "frmSplash.frx":A0F9D
      textCB          =   "frmSplash.frx":A0FB5
      textRB          =   "frmSplash.frx":A0FCD
      colorBack       =   "frmSplash.frx":A0FE5
      colorIntern     =   "frmSplash.frx":A100F
      colorMO         =   "frmSplash.frx":A1039
      colorFocus      =   "frmSplash.frx":A1063
      colorDisabled   =   "frmSplash.frx":A108D
      colorPressed    =   "frmSplash.frx":A10B7
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   5415
      Left            =   390
      TabIndex        =   25
      Top             =   4530
      Width           =   4305
      _Version        =   524298
      _ExtentX        =   7594
      _ExtentY        =   9551
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      BackColorContainer=   11131362
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A10E1
      textLT          =   "frmSplash.frx":A10F9
      textCT          =   "frmSplash.frx":A1111
      textRT          =   "frmSplash.frx":A1129
      textLM          =   "frmSplash.frx":A1141
      textRM          =   "frmSplash.frx":A1159
      textLB          =   "frmSplash.frx":A1171
      textCB          =   "frmSplash.frx":A1189
      textRB          =   "frmSplash.frx":A11A1
      colorBack       =   "frmSplash.frx":A11B9
      colorIntern     =   "frmSplash.frx":A11E3
      colorMO         =   "frmSplash.frx":A120D
      colorFocus      =   "frmSplash.frx":A1237
      colorDisabled   =   "frmSplash.frx":A1261
      colorPressed    =   "frmSplash.frx":A128B
      Style           =   3
   End
   Begin BTNENHLib4.BtnEnh BtnEnh2 
      Height          =   6075
      Left            =   4860
      TabIndex        =   26
      Top             =   3330
      Width           =   4965
      _Version        =   524298
      _ExtentX        =   8758
      _ExtentY        =   10716
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      CornerFactor    =   20
      BackColorContainer=   11131362
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A12B5
      textLT          =   "frmSplash.frx":A12CD
      textCT          =   "frmSplash.frx":A12E5
      textRT          =   "frmSplash.frx":A12FD
      textLM          =   "frmSplash.frx":A1315
      textRM          =   "frmSplash.frx":A132D
      textLB          =   "frmSplash.frx":A1345
      textCB          =   "frmSplash.frx":A135D
      textRB          =   "frmSplash.frx":A1375
      colorBack       =   "frmSplash.frx":A138D
      colorIntern     =   "frmSplash.frx":A13B7
      colorMO         =   "frmSplash.frx":A13E1
      colorFocus      =   "frmSplash.frx":A140B
      colorDisabled   =   "frmSplash.frx":A1435
      colorPressed    =   "frmSplash.frx":A145F
      Style           =   3
   End
   Begin BTNENHLib4.BtnEnh cmdChange 
      Height          =   915
      Left            =   4860
      TabIndex        =   32
      Top             =   2340
      Visible         =   0   'False
      Width           =   4965
      _Version        =   524298
      _ExtentX        =   8758
      _ExtentY        =   1614
      _StockProps     =   66
      Caption         =   "Change"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21
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
         Size            =   20.25
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
      BackColorContainer=   11131362
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A1489
      textLT          =   "frmSplash.frx":A14F5
      textCT          =   "frmSplash.frx":A150D
      textRT          =   "frmSplash.frx":A1525
      textLM          =   "frmSplash.frx":A153D
      textRM          =   "frmSplash.frx":A1555
      textLB          =   "frmSplash.frx":A156D
      textCB          =   "frmSplash.frx":A1585
      textRB          =   "frmSplash.frx":A159D
      colorBack       =   "frmSplash.frx":A15B5
      colorIntern     =   "frmSplash.frx":A15DF
      colorMO         =   "frmSplash.frx":A1609
      colorFocus      =   "frmSplash.frx":A1633
      colorDisabled   =   "frmSplash.frx":A165D
      colorPressed    =   "frmSplash.frx":A1687
      Style           =   3
   End
   Begin BTNENHLib4.BtnEnh fmClock 
      Height          =   1635
      Left            =   5130
      TabIndex        =   33
      Top             =   9480
      Visible         =   0   'False
      Width           =   4485
      _Version        =   524298
      _ExtentX        =   7911
      _ExtentY        =   2884
      _StockProps     =   66
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
         Size            =   9.75
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
      BackColorContainer=   5879007
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplash.frx":A16B1
      textLT          =   "frmSplash.frx":A16C9
      textCT          =   "frmSplash.frx":A16E1
      textRT          =   "frmSplash.frx":A16F9
      textLM          =   "frmSplash.frx":A1711
      textRM          =   "frmSplash.frx":A1729
      textLB          =   "frmSplash.frx":A1741
      textCB          =   "frmSplash.frx":A1759
      textRB          =   "frmSplash.frx":A1771
      colorBack       =   "frmSplash.frx":A1789
      colorIntern     =   "frmSplash.frx":A17B3
      colorMO         =   "frmSplash.frx":A17DD
      colorFocus      =   "frmSplash.frx":A1807
      colorDisabled   =   "frmSplash.frx":A1831
      colorPressed    =   "frmSplash.frx":A185B
      Style           =   3
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Expires on:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   9930
      TabIndex        =   39
      Top             =   10710
      Width           =   1245
   End
   Begin VB.Label lblenddate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9930
      TabIndex        =   38
      Top             =   10980
      Width           =   5325
   End
   Begin VB.Label lblstartdate 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7260
      TabIndex        =   37
      Top             =   11160
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblServer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   10170
      TabIndex        =   35
      Top             =   90
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   210
      TabIndex        =   34
      Top             =   10980
      Width           =   3855
   End
   Begin MSForms.ComboBox cmbUsers 
      Height          =   510
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   3525
      BackColor       =   8505299
      ForeColor       =   16777215
      DisplayStyle    =   7
      Size            =   "6218;900"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblWorkstation 
      Height          =   315
      Left            =   210
      TabIndex        =   23
      Top             =   10020
      Width           =   2145
      ForeColor       =   16777215
      BackColor       =   12632256
      VariousPropertyBits=   8388627
      Caption         =   "1"
      Size            =   "3784;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   13770
      TabIndex        =   21
      Top             =   420
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   10335
      Width           =   4485
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   13770
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   10650
      Width           =   2145
   End
   Begin MSForms.Image newBack 
      Height          =   1005
      Left            =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "1764;1764"
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbUsers_Change()
    On Error Resume Next
    If cmbUsers.Text = "" Then Exit Sub
    If Trim(Server.SQL_Name) <> "" Then
        ActiveReadServer "Select * from Users where user_no=" & Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2))
        If rs.RecordCount > 0 Then
            UserRecord.User_Number = rs.Fields("User_No")
            UserRecord.Com_Calc = Val(rs.Fields("Com_Calc") & "")
            UserRecord.Bar_Cash = Val(rs.Fields("Bar_Cash") & "")
            UserRecord.Name = rs.Fields("User_Name")
            UserRecord.Password = rs.Fields("User_Password")
            UserRecord.FirstName = rs.Fields("First_Name")
            UserRecord.LastName = rs.Fields("Last_Name")
            UserRecord.Wage = Val(rs.Fields("Wage") & "")
            UserRecord.Comm1 = Val(rs.Fields("Comm1") & "")
            UserRecord.Comm2 = Val(rs.Fields("Comm1") & "")
            UserRecord.Drawer_No = Val(rs.Fields("Drawer_No") & "")
            If UserRecord.Drawer_No = 0 Then UserRecord.Drawer_No = 1
            If IsNull(rs.Fields("Logged_In")) Then
                UserRecord.Logged_in = False
            Else
                UserRecord.Logged_in = rs.Fields("Logged_In")
            End If
            If IsNull(rs.Fields("All_Tables")) Then
                UserRecord.All_Tables = False
            Else
                UserRecord.All_Tables = rs.Fields("All_Tables")
            End If
            If IsNull(rs.Fields("Ua_Checkin")) Then
                UserRecord.Checkin = False
            Else
                UserRecord.Checkin = rs.Fields("Ua_Checkin")
            End If
            If IsNull(rs.Fields("Ua_Checkout")) Then
                UserRecord.Checkout = False
            Else
                UserRecord.Checkout = rs.Fields("Ua_Checkout")
            End If
            If IsNull(rs.Fields("Ua_Guests")) Then
                UserRecord.Guests = False
            Else
                UserRecord.Guests = rs.Fields("Ua_Guests")
            End If
            If IsNull(rs.Fields("Ua_Rooms")) Then
                UserRecord.Rooms = False
            Else
                UserRecord.Rooms = rs.Fields("Ua_Rooms")
            End If
            If IsNull(rs.Fields("Ua_Inventory")) Then
                UserRecord.Inventory = False
            Else
                UserRecord.Inventory = rs.Fields("Ua_Inventory")
            End If
            If IsNull(rs.Fields("Ua_Sales")) Then
                UserRecord.Sales = False
            Else
                UserRecord.Sales = rs.Fields("Ua_Sales")
            End If
            If IsNull(rs.Fields("Ua_Reports")) Then
                UserRecord.Reports = False
            Else
                UserRecord.Reports = rs.Fields("Ua_Reports")
            End If
            If IsNull(rs.Fields("Ua_Settings")) Then
                UserRecord.Settings = False
            Else
                UserRecord.Settings = rs.Fields("Ua_Settings")
            End If
            If IsNull(rs.Fields("Ua_Users")) Then
                UserRecord.Users = False
            Else
                UserRecord.Users = rs.Fields("Ua_Users")
            End If
            If IsNull(rs.Fields("Ua_Reservations")) Then
                UserRecord.Reservations = False
            Else
                UserRecord.Reservations = rs.Fields("Ua_Reservations")
            End If
            UserRecord.uType = rs.Fields("User_Type")
            If IsNull(rs.Fields("Cash_Sales")) Then
                UserRecord.Card_Sales = False
            Else
                UserRecord.Cash_Sales = rs.Fields("Cash_Sales")
            End If
            If IsNull(rs.Fields("Cheque_Sales")) Then
                UserRecord.Cheque_Sales = False
            Else
                UserRecord.Cheque_Sales = rs.Fields("Cheque_Sales")
            End If
            If IsNull(rs.Fields("Card_Sales")) Then
                UserRecord.Card_Sales = False
            Else
                UserRecord.Card_Sales = rs.Fields("Card_Sales")
            End If
            If IsNull(rs.Fields("Charge_Sales")) Then
                UserRecord.Charge_Sales = False
            Else
                UserRecord.Charge_Sales = rs.Fields("Charge_Sales")
            End If
            If IsNull(rs.Fields("Loyalty_Sales")) Then
                UserRecord.Loyalty_Sales = False
            Else
                UserRecord.Loyalty_Sales = rs.Fields("Loyalty_Sales")
            End If
            If IsNull(rs.Fields("Over_Tender")) Then
                UserRecord.Over_Tender = False
            Else
                UserRecord.Over_Tender = rs.Fields("Over_Tender")
            End If
            If IsNull(rs.Fields("Disc_Perc")) Then
                UserRecord.Disc_Perc = False
            Else
                UserRecord.Disc_Perc = rs.Fields("Disc_Perc")
            End If
            If IsNull(rs.Fields("Item_Corrects")) Then
                UserRecord.Item_Corrects = False
            Else
                UserRecord.Item_Corrects = rs.Fields("Item_Corrects")
            End If
            If IsNull(rs.Fields("Voids")) Then
                UserRecord.Voids = False
            Else
                UserRecord.Voids = rs.Fields("Voids")
            End If
            If IsNull(rs.Fields("Returns")) Then
                UserRecord.Returns = False
            Else
                UserRecord.Returns = rs.Fields("Returns")
            End If
            If IsNull(rs.Fields("Ullages")) Then
                UserRecord.Ullages = False
            Else
                UserRecord.Ullages = rs.Fields("Ullages")
            End If
            If IsNull(rs.Fields("Disc_Amt")) Then
                UserRecord.Disc_Amt = False
            Else
                UserRecord.Disc_Amt = rs.Fields("Disc_Amt")
            End If
            If IsNull(rs.Fields("Payouts")) Then
                UserRecord.Payouts = False
            Else
                UserRecord.Payouts = rs.Fields("Payouts")
            End If
            If IsNull(rs.Fields("Pickups")) Then
                UserRecord.Pickups = False
            Else
                UserRecord.Pickups = rs.Fields("Pickups")
            End If
            If IsNull(rs.Fields("App_Exit")) Then
                UserRecord.App_Exit = False
            Else
                UserRecord.App_Exit = rs.Fields("App_Exit")
            End If
            If IsNull(rs.Fields("Loans")) Then
                UserRecord.Loans = False
            Else
                UserRecord.Loans = rs.Fields("Loans")
            End If
            If IsNull(rs.Fields("Receive_Acc")) Then
                UserRecord.Receive_Acc = False
            Else
                UserRecord.Receive_Acc = rs.Fields("Receive_Acc")
            End If
            If IsNull(rs.Fields("Split_Tenders")) Then
                UserRecord.Split_Tender = False
            Else
                UserRecord.Split_Tender = rs.Fields("Split_Tenders")
            End If
            If IsNull(rs.Fields("Buffer_Print")) Then
                UserRecord.Buffer_Print = False
            Else
                UserRecord.Buffer_Print = rs.Fields("Buffer_Print")
            End If
            If IsNull(rs.Fields("Reprint")) Then
                UserRecord.Reprint = False
            Else
                UserRecord.Reprint = rs.Fields("Reprint")
            End If
            If IsNull(rs.Fields("Trans_Store")) Then
                UserRecord.Trans_Store = False
            Else
                UserRecord.Trans_Store = rs.Fields("Trans_Store")
            End If
            If IsNull(rs.Fields("Trans_Clear")) Then
                UserRecord.Trans_Clear = False
            Else
                UserRecord.Trans_Clear = rs.Fields("Trans_Clear")
            End If
            If IsNull(rs.Fields("Transfer")) Then
                UserRecord.Transfers = False
            Else
                UserRecord.Transfers = rs.Fields("Transfer")
            End If
            If IsNull(rs.Fields("Over_Tender")) Then
                UserRecord.Over_Tender = False
            Else
                UserRecord.Over_Tender = rs.Fields("Over_Tender")
            End If
            If IsNull(rs.Fields("OverRide")) Then
                UserRecord.Overides = False
            Else
                UserRecord.Overides = rs.Fields("OverRide")
            End If
            If IsNull(rs.Fields("Search")) Then
                UserRecord.Search = False
            Else
                UserRecord.Search = rs.Fields("Search")
            End If
            If IsNull(rs.Fields("Total_Clear")) Then
                UserRecord.Total_Clear = False
            Else
                UserRecord.Total_Clear = rs.Fields("Total_Clear")
            End If
            If IsNull(rs.Fields("Draw_Cash")) Then
                UserRecord.Draw_Cash = False
            Else
                UserRecord.Draw_Cash = rs.Fields("Draw_Cash")
            End If
            If IsNull(rs.Fields("Draw_Card")) Then
                UserRecord.Draw_Card = False
            Else
                UserRecord.Draw_Card = rs.Fields("Draw_Card")
            End If
            If IsNull(rs.Fields("Draw_Cheque")) Then
                UserRecord.Draw_Cheque = False
            Else
                UserRecord.Draw_Cheque = rs.Fields("Draw_Cheque")
            End If
            If IsNull(rs.Fields("Draw_Charge")) Then
                UserRecord.Draw_Charge = False
            Else
                UserRecord.Draw_Charge = rs.Fields("Draw_Charge")
            End If
            If IsNull(rs.Fields("No_Sales")) Then
                UserRecord.No_Sales = False
            Else
                UserRecord.No_Sales = rs.Fields("No_Sales")
            End If
            If IsNull(rs.Fields("Draw_Loyalty")) Then
                UserRecord.Draw_Loyalty = False
            Else
                UserRecord.Draw_Loyalty = rs.Fields("Draw_Loyalty")
            End If
            
            If IsNull(rs.Fields("Quotes")) Then
                UserRecord.Quotes = False
            Else
                UserRecord.Quotes = rs.Fields("Quotes")
            End If
            
            'Kotie 22-03-2013 13:10
            If IsNull(rs.Fields("Owner_transfer")) Then
                UserRecord.Quotes = False
            Else
                UserRecord.Owner_Transfer = rs.Fields("Owner_transfer")
            End If
            
        End If
        rs.Close
    End If
    If Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2)) = 0 Then
        UserRecord.User_Number = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Number", Default:=Trim(gblApp_Name)))
        UserRecord.Name = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Name", Default:=Trim(gblApp_Name)))
        UserRecord.Password = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Password", Default:=Trim(gblApp_Name)))
    End If
    
    txtPassword.Text = ""
    On Error GoTo 0
End Sub

Private Sub cmdClose_Click()
      End
End Sub

Private Sub cmbUsers_GotFocus()
    If System_Access = 1 Then
        txtPassword.SetFocus
    End If
End Sub

Private Sub cmdErr_Click()
    Timer4.Enabled = False
    cmdErr.BackColor = &H30A8CF
    cmdErr.Visible = False
    txtPassword.Tag = ""
    cmdInput(14).SetFocus
End Sub
Private Sub cmdInput_Click(Index As Integer)



    cmdChange.Visible = False
    Select Case cmdInput(Index).Caption
    Case "0" To "9"
    
    Case Else
        send_data_steam_keylog (Me.Name & " - " & cmdInput(Index).Caption)
    End Select
    
    Select Case cmdInput(Index).Caption
        Case "Stock Takes"
            If cmbUsers.Text = "" Then
                Load frmError
                frmError.lblCap.Caption = "Please select a user name before entering your password."
                frmError.Show vbModal
                txtPassword.Text = ""
                cmbUsers.SetFocus
                Exit Sub
            End If
            If txtPassword.Text = "" Then
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                Exit Sub
            End If
           
            ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
            If rs.RecordCount > 0 Then
                newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.Close
                Savepassword = txtPassword.Text
                cmbUsers.Text = newText
                fmUser.Caption = cmbUsers.Text
                DoEvents
                txtPassword.Text = Savepassword
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                rs.Close
            End If
            If Trim(UserRecord.Password) = Trim(txtPassword) Then
                If UserRecord.Inventory = False Then
                    Load frmError
                    frmError.lblCap.Caption = "You do not have Access Rights to do Stock Takes."
                    frmError.Show vbModal
                    txtPassword.Text = ""
                    On Error Resume Next
                    txtPassword.SetFocus
                    On Error GoTo 0
                    Exit Sub
                Else
                    frmStockTake.Show vbModal
                End If
            End If
        Case "0" To "9"
            txtPassword.SetFocus
            SendKeys cmdInput(Index).Caption
            DoEvents
        Case "Ok"
            txtPassword.SetFocus
            SendKeys "{ENTER}"
        Case "CL"
            Timer4.Enabled = False
            cmdErr.BackColor = &H30A8CF
            cmdErr.Visible = False
            txtPassword.Tag = ""
            txtPassword.SetFocus
            txtPassword.Text = ""
        Case "Exit"
            If cmbUsers.Text = "" Then
                Load frmError
                frmError.lblCap.Caption = "Please select a user name before entering your password."
                frmError.Show vbModal
                txtPassword.Text = ""
                On Error Resume Next
                cmbUsers.SetFocus
                On Error GoTo 0
                Exit Sub
            End If
            If txtPassword.Text = "" Then
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                Exit Sub
            End If
            ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
            If rs.RecordCount > 0 Then
                newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.Close
                Savepassword = txtPassword.Text
                cmbUsers.Text = newText
                fmUser.Caption = cmbUsers.Text
                DoEvents
                txtPassword.Text = Savepassword
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                rs.Close
            End If
            If Trim(UserRecord.Password) = Trim(txtPassword) Then
                If UserRecord.App_Exit = False Then
                    Load frmError
                    frmError.lblCap.Caption = "You do not have Access Rights to Exit."
                    frmError.Show vbModal
                    txtPassword.Text = ""
                    On Error Resume Next
                    cmbUsers.SetFocus
                    On Error GoTo 0
                    Exit Sub
                Else
                    cnnMain.Close
                    End
                End If
            End If
        Case "Stock Requests"
            If System_Access = 1 And txtPassword.Text <> "" Then
                ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
                If rs.RecordCount > 0 Then
                    newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    rs.Close
                    Savepassword = txtPassword.Text
                    cmbUsers.Text = newText
                    fmUser.Caption = cmbUsers.Text
                    DoEvents
                    txtPassword.Text = Savepassword
                Else
                    rs.Close
                End If
            End If
            Timer4.Enabled = False
            cmdErr.BackColor = &H30A8CF
            cmdErr.Visible = False
            txtPassword.Tag = ""
            KeyCode = 0
            If txtPassword.Text = "" Then
                txtPassword.Tag = "Stock Requests"
                cmdErr.Caption = "Please Sign on to Access Stock Requests"
                cmdErr.Visible = True
                Timer4.Enabled = True
                txtPassword.SetFocus
                Exit Sub
            End If
            If cmbUsers.Text = "" Then
                Load frmError
                frmError.lblCap.Caption = "Please select a user name before entering your password."
                frmError.Show vbModal
                txtPassword.Text = ""
                cmbUsers.SetFocus
                Exit Sub
            End If
            If Trim(UserRecord.Password) = Trim(txtPassword) Then
                ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                If rs.RecordCount > 0 Then
                    If rs.Fields("Function_Key") = 4 Then
                        Load frmError
                        frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Making Stock Requests"
                        frmError.Show vbModal
                        txtPassword.SetFocus
                        txtPassword.Text = ""
                        rs.Close
                        Exit Sub
                    End If
                Else
                    Load frmError
                    frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Making Stock Requests"
                    frmError.Show vbModal
                    txtPassword.SetFocus
                    txtPassword.Text = ""
                    rs.Close
                    Exit Sub
                End If
                rs.Close
                UserRecord.LastUser = Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2))
                Select Case UserRecord.Inventory
                    Case False
                        Load frmError
                        frmError.lblCap.Caption = "You do not have access rights to make Stock Requests."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                        Exit Sub
                    Case True
                        Screen.MousePointer = 11
                        KeyCode = 0
                        GlobalMode = TillMode.StartMode
                        frmRequest.Show vbModal
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Stock Requests')"
                        DoEvents
                        Screen.MousePointer = 0
                End Select
                DoEvents
                SaveSetting Trim(gblApp_Name), "User", "Last_User", UserRecord.LastUser
                If Trim(Server.SQL_Name) = "" Then
                    frmSettings.Show vbModal
                End If
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                Exit Sub
            End If
        Case "Reservations"
            If System_Access = 1 And txtPassword.Text <> "" Then
                ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
                If rs.RecordCount > 0 Then
                    newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    rs.Close
                    Savepassword = txtPassword.Text
                    cmbUsers.Text = newText
                    fmUser.Caption = cmbUsers.Text
                    DoEvents
                    txtPassword.Text = Savepassword
                Else
                    rs.Close
                End If
            End If
            Timer4.Enabled = False
            cmdErr.BackColor = &H30A8CF
            cmdErr.Visible = False
            txtPassword.Tag = ""
            KeyCode = 0
            If txtPassword.Text = "" Then
                txtPassword.Tag = "Stock Requests"
                cmdErr.Caption = "Please Sign on to Access Reservations"
                cmdErr.Visible = True
                Timer4.Enabled = True
                txtPassword.SetFocus
                Exit Sub
            End If
            If cmbUsers.Text = "" Then
                Load frmError
                frmError.lblCap.Caption = "Please select a user name before entering your password."
                frmError.Show vbModal
                txtPassword.Text = ""
                cmbUsers.SetFocus
                Exit Sub
            End If
            If Trim(UserRecord.Password) = Trim(txtPassword) Then
                ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                If rs.RecordCount > 0 Then
                    If rs.Fields("Function_Key") = 4 Then
                        Load frmError
                        frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Making Reservations"
                        frmError.Show vbModal
                        txtPassword.SetFocus
                        txtPassword.Text = ""
                        rs.Close
                        Exit Sub
                    End If
                Else
                    Load frmError
                    frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Reservations"
                    frmError.Show vbModal
                    txtPassword.SetFocus
                    txtPassword.Text = ""
                    rs.Close
                    Exit Sub
                End If
                rs.Close
                UserRecord.LastUser = Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2))
                Select Case UserRecord.Reports
                    Case False
                        Load frmError
                        frmError.lblCap.Caption = "You do not have access rights to make Reservations."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                        Exit Sub
                    Case True
                        Screen.MousePointer = 11
                        KeyCode = 0
                        GlobalMode = TillMode.StartMode
                        frmRestRes.Show
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Reservations')"
                        DoEvents
                        Me.Hide
                        Screen.MousePointer = 0
                End Select
                DoEvents
                SaveSetting Trim(gblApp_Name), "User", "Last_User", UserRecord.LastUser
                If Trim(Server.SQL_Name) = "" Then
                    frmSettings.Show vbModal
                End If
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                Exit Sub
            End If
        Case "Cashup's"
            If System_Access = 1 And txtPassword.Text <> "" Then
                ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
                If rs.RecordCount > 0 Then
                    newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    rs.Close
                    Savepassword = txtPassword.Text
                    cmbUsers.Text = newText
                    fmUser.Caption = cmbUsers.Text
                    DoEvents
                    txtPassword.Text = Savepassword
                Else
                    rs.Close
                End If
            End If
            Timer4.Enabled = False
            cmdErr.BackColor = &H30A8CF
            cmdErr.Visible = False
            txtPassword.Tag = ""
            KeyCode = 0
            If txtPassword.Text = "" Then
                txtPassword.Tag = "Cash-up's"
                cmdErr.Caption = "Please Sign on to Access Cash-up's"
                cmdErr.Visible = True
                Timer4.Enabled = True
                txtPassword.SetFocus
                Exit Sub
            End If
            If cmbUsers.Text = "" Then
                Load frmError
                frmError.lblCap.Caption = "Please enter your password again."
                frmError.Show vbModal
                txtPassword.Text = ""
                cmbUsers.SetFocus
                Exit Sub
            End If
            If Trim(UserRecord.Password) = Trim(txtPassword) Then
                UserRecord.LastUser = Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2))
                Select Case UserRecord.Reports
                    Case False
                        ActiveReadServer "Select Cashup_No from Counters where Finalized= 0 and User_No=" & UserRecord.User_Number
                        If rs.RecordCount > 0 Then
                            Cashup_No = rs.Fields("Cashup_No")
                            rs.Close
                            
                            ActiveReadServer "Select count(Table_No) as TableCount from Table_Listing_View where User_No = " & UserRecord.User_Number
                            If rs.Fields("TableCount") > 0 Then
                                Timer4.Enabled = True
                                Select Case rs.Fields("TableCount")
                                    Case 1
                                        cmdErr.Caption = "You still have an Open Table."
                                    Case Else
                                        cmdErr.Caption = "You still have " & rs.Fields("TableCount") & " Open Tables."
                                End Select
                                cmdErr.Visible = True
                                txtPassword.Text = ""
                                rs.Close
                                Exit Sub
                            End If
                            rs.Close
                            
                            ActiveReadServer "Select count(Tab_No) as TabCount from Tab_Listing_View where User_No = " & UserRecord.User_Number
                            If rs.Fields("TabCount") > 0 Then
                                Timer4.Enabled = True
                                Select Case rs.Fields("TabCount")
                                    Case 1
                                        cmdErr.Caption = "You still have an Open Tab."
                                    Case Else
                                        cmdErr.Caption = "You still have " & rs.Fields("TabCount") & " Open Tabs."
                                End Select
                                cmdErr.Visible = True
                                txtPassword.Text = ""
                                rs.Close
                                Exit Sub
                            End If
                            rs.Close
                            Screen.MousePointer = 11
                            
                            Load frmCapture
                            frmCapture.Tag = -10
                            frmCapture.lblHeading.Tag = Cashup_No
                            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Cash-up Capture')"
                            Screen.MousePointer = 0
                            frmCapture.Show vbModal
                            frmCapture.lblHeading.Tag = ""
                            Exit Sub
                        Else
                            rs.Close
                            Load frmError
                            frmError.lblCap.Caption = "You do not have access rights to use Cash-up's."
                            frmError.Show vbModal
                            txtPassword.Text = ""
                            txtPassword.SetFocus
                            Exit Sub
                        End If
                    Case True
                        Screen.MousePointer = 11
                        KeyCode = 0
                        GlobalMode = TillMode.CashupMode
                        frmTillReport.Show
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Cash-up`s')"
                        DoEvents
                        Me.Hide
                        Screen.MousePointer = 0
                End Select
                DoEvents
                SaveSetting Trim(gblApp_Name), "User", "Last_User", UserRecord.LastUser
                If Trim(Server.SQL_Name) = "" Then
                    frmSettings.Show vbModal
                End If
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                Exit Sub
            End If
        Case "Clock In"
            If System_Access = 1 And txtPassword.Text <> "" Then
                ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
                If rs.RecordCount > 0 Then
                    newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    rs.Close
                    Savepassword = txtPassword.Text
                    cmbUsers.Text = newText
                    fmUser.Caption = cmbUsers.Text
                    DoEvents
                    txtPassword.Text = Savepassword
                Else
                    rs.Close
                End If
            End If
            Timer4.Enabled = False
            cmdErr.BackColor = &H30A8CF
            cmdErr.Visible = False
            txtPassword.Tag = ""
            If txtPassword.Text = "" Then
                txtPassword.Tag = "Clock In"
                cmdErr.Caption = "Please Sign on to Clock In"
                cmdErr.Visible = True
                Timer4.Enabled = True
                txtPassword.SetFocus
                Exit Sub
            End If
            If Trim(UserRecord.Password) = Trim(txtPassword) Then
                 ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                        "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                        
                
                If rs.RecordCount > 0 Then
                    If rs.Fields("Function_Key") = 3 Then
                        Load frmError
                        frmError.lblCap.Caption = "You have already Clocked In. Enter password and click OK or Clock Out."
                        frmError.Show vbModal
                        txtPassword.SetFocus
                        txtPassword.Text = ""
                        rs.Close
                        Exit Sub
                    End If
                End If
                rs.Close
                
               
                If Devices.TwoDrawer = 1 And UserRecord.uType = 4 Then
                    Load frmChooseRes
                    frmSplash.Tag = "Not Now"
                    frmChooseRes.Caption = "Please Select a Drawer for this Shift."
                    DoEvents
                    frmChooseRes.Show vbModal
                    Select Case frmSplash.cmbUsers.Tag
                        Case "Drawer One"
                            Drawer = 1
                        Case "Drawer Two"
                            Drawer = 2
                        Case Else
                            frmError.lblCap.Caption = "You have to Select a Drawer before you will be able to Clock In"
                            frmError.Show vbModal
                            txtPassword.SetFocus
                            txtPassword.Text = ""
                            Exit Sub
                    End Select
                Else
                    Drawer = 1
                End If
                fmClock.Visible = True
                fmClock.Caption = UserRecord.Name & " Clocked in at " & Format(Time, "HH:MM")
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),3," & Workstation_No & ")"
                ActiveUpdateServer "Update Users set Workstation_No=" & Workstation_No & ",Drawer_No=" & Drawer & ",Clocked_In = 1 where User_No = " & UserRecord.User_Number
                Print_Clock_In
                txtPassword.Text = ""
                txtPassword.SetFocus
            '*****************************************************
            
            
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
            End If
        Case "Clock Out"
            If System_Access = 1 And txtPassword.Text <> "" Then
                ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
                If rs.RecordCount > 0 Then
                    newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    rs.Close
                    Savepassword = txtPassword.Text
                    cmbUsers.Text = newText
                    fmUser.Caption = cmbUsers.Text
                    DoEvents
                    txtPassword.Text = Savepassword
                Else
                    rs.Close
                End If
            End If
            Timer4.Enabled = False
            cmdErr.BackColor = &H30A8CF
            cmdErr.Visible = False
            txtPassword.Tag = ""
            If txtPassword.Text = "" Then
                txtPassword.Tag = "Clock Out"
                cmdErr.Caption = "Please Sign on to Clock Out"
                cmdErr.Visible = True
                Timer4.Enabled = True
                txtPassword.SetFocus
                Exit Sub
            End If
            ActiveReadServer "Select count(Table_No) as TableCount from Table_Listing_View where User_No = " & UserRecord.User_Number
            Select Case rs.Fields("TableCount")
                Case 0
                Case 1
                    Load frmError
                    frmError.lblCap.Caption = "You still have an Open Table. You cannot Clock Out while you have an Open Table."
                    frmError.Show vbModal
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                    rs.Close
                    Exit Sub
                Case Else
                    Load frmError
                    cmdErr.Caption = "You still have " & rs.Fields("TableCount") & " Open Tables. You cannot Clock Out while you have an Open Table."""
                    frmError.Show vbModal
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                    rs.Close
                    Exit Sub
            End Select
            rs.Close
            
'            ActiveReadServer "Select * from Cash_ups where Finalized = 0 and User_No = " & UserRecord.User_Number
'            If rs.RecordCount > 0 Then
'                rs.Close
'                Load frmError
'                frmError.lblCap.Caption = "You cannot Clock Out while you have an Outstanding Cashup."
'                frmError.Show vbModal
'                txtPassword.Text = ""
'                txtPassword.SetFocus
'                Exit Sub


            ActiveReadServer "Select * from Cash_ups where Finalized = 0 and User_No = " & UserRecord.User_Number
          
            If rs.RecordCount > 0 Then
                rs.Close
                Load frmQuestion2
                
                frmQuestion2.Show vbModal
                If BtnEnh2.Tag = "Dothecapture" Then
                 Call Dothecapture
                End If
                If BtnEnh2.Tag = "Dotheclockout" Then
                Call Dotheclockout
                End If

                Exit Sub


            End If
            rs.Close
            If Trim(UserRecord.Password) = Trim(txtPassword) Then
                ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                If rs.RecordCount > 0 Then
                    If rs.Fields("Function_Key") = 4 Then
                        Load frmError
                        frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Clocking Out"
                        frmError.Show vbModal
                        txtPassword.SetFocus
                        txtPassword.Text = ""
                        rs.Close
                        Exit Sub
                    End If
                Else
                    Load frmError
                    frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Clocking Out"
                    frmError.Show vbModal
                    txtPassword.SetFocus
                    txtPassword.Text = ""
                    rs.Close
                    Exit Sub
                End If
                rs.Close
                fmClock.Caption = UserRecord.Name & " Clocked Out at " & Format(Time, "HH:MM")
                Print_Clock_Out
                fmClock.Visible = True
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),4," & Workstation_No & ")"
                ActiveUpdateServer "Update Users set Drawer_No = 0,Clocked_In = 0 where User_No = " & UserRecord.User_Number
                txtPassword.Text = ""
                txtPassword.SetFocus
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
            End If
    End Select
End Sub
Private Sub Dothecapture()
  If System_Access = 1 And txtPassword.Text <> "" Then
                ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
                If rs.RecordCount > 0 Then
                    newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    rs.Close
                    Savepassword = txtPassword.Text
                    cmbUsers.Text = newText
                    fmUser.Caption = cmbUsers.Text
                    DoEvents
                    txtPassword.Text = Savepassword
                    GoTo Continue1
                 End If
                 End If
                 Exit Sub
                 
                 
Continue1:

 If Trim(UserRecord.Password) = Trim(txtPassword) Then
                ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                If rs.RecordCount > 0 Then
                    If rs.Fields("Function_Key") = 4 Then
                        Load frmError
                        frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Clocking Out"
                        frmError.Show vbModal
                        txtPassword.SetFocus
                        txtPassword.Text = ""
                        'rs.Close
                        Exit Sub
                    End If
                Else
                    Load frmError
                    frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Clocking Out"
                    frmError.Show vbModal
                    txtPassword.SetFocus
                    txtPassword.Text = ""
                    'rs.Close
                    Exit Sub
                End If
                Senttofinalize = True
cmdInput_Click (16)

fmClock.Caption = UserRecord.Name & " Clocked Out at " & Format(Time, "HH:MM")
                Print_Clock_Out
                fmClock.Visible = True
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),4," & Workstation_No & ")"
                ActiveUpdateServer "Update Users set Drawer_No = 0,Clocked_In = 0 where User_No = " & UserRecord.User_Number
                txtPassword.Text = ""
                If txtPassword.Visible Then
                txtPassword.SetFocus
                End If
Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                If rs.State = 1 Then rs.Close
                Exit Sub
            End If
    
            If rs.State = 1 Then rs.Close
End Sub
Private Sub Dotheclockout()
                 
                 If System_Access = 1 And txtPassword.Text <> "" Then
                ActiveReadServer "Select User_No,User_Name from Users where User_Password = " & txtPassword.Text
                If rs.RecordCount > 0 Then
                    newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    rs.Close
                    Savepassword = txtPassword.Text
                    cmbUsers.Text = newText
                    fmUser.Caption = cmbUsers.Text
                    DoEvents
                    txtPassword.Text = Savepassword
                    GoTo Continue
                 End If
                 End If
                 Exit Sub
                 
                 
Continue:
                 
       If Trim(Savepassword) = Trim(txtPassword) Then
                ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                If rs.RecordCount > 0 Then
                    If rs.Fields("Function_Key") = 4 Then
                        Load frmError
                        frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Clocking Out"
                        frmError.Show vbModal
                        txtPassword.SetFocus
                        txtPassword.Text = ""
                        Exit Sub
                    End If
                Else
                    Load frmError
                    frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Clocking Out"
                    frmError.Show vbModal
                    txtPassword.SetFocus
                    txtPassword.Text = ""
                    'rs.Close
                    Exit Sub
                End If
                'rs.Close
                fmClock.Caption = UserRecord.Name & " Clocked Out at " & Format(Time, "HH:MM")
                Print_Clock_Out
                fmClock.Visible = True
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),4," & Workstation_No & ")"
                ActiveUpdateServer "Update Users set Drawer_No = 0,Clocked_In = 0 where User_No = " & UserRecord.User_Number
                txtPassword.Text = ""
                txtPassword.SetFocus
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                If rs.State = 1 Then rs.Close
                Exit Sub
            End If
    
            If rs.State = 1 Then rs.Close
End Sub





Private Sub Form_Activate()
    SafetyCode "Specials"
    On Error Resume Next
    txtPassword.Text = ""
    lblServer.Caption = "Server: " & Trim(Server.SQL_Name) & " - Database: " & Trim(Server.SQL_Database)
    If Me.Height < 10000 And newBack.Visible = False Then
        picAm.Visible = False
        newBack.Visible = True
        On Error Resume Next
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.782
            Me.Controls(i).Left = Me.Controls(i).Left * 0.782
            Me.Controls(i).Height = Me.Controls(i).Height * 0.79
            Me.Controls(i).top = Me.Controls(i).top * 0.79
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
        newBack.Width = Screen.Width
        newBack.Height = Screen.Height
        On Error GoTo 0
    End If
    If frmSplash.Tag = "Not Now" Then
        frmSplash.Tag = ""
        Exit Sub
    End If
    txtPassword.Text = ""
    newBack.Width = Screen.Width
    newBack.Height = Screen.Height
    frmSplash.Tag = ""
    lblWorkstation.Caption = Workstation_No & " - " & Workstation_Name
    Label1.Caption = "Copyright 2012 HeroPOS" & Chr(169)
    Timer4.Enabled = False
    cmdErr.BackColor = &H30A8CF
    cmdErr.Visible = False
    txtPassword.Tag = ""
    Screen.MousePointer = 11
    If cmbUsers.Text <> "" Then txtPassword.SetFocus
    txtPassword.Text = ""
    lblTime.Caption = Time
    DoEvents
    lblVersion.Caption = "Version: " & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
    lblDate.Caption = Format(Date, "YYYY-MM-DD")
    Label2.Caption = "Licensed to  " & Trim(Server.SQL_Database)
    'Look for default user
    Select Case GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Name", Default:="StartUp")
        Case "StartUp"
            'If user do not exists add a default user
            SaveSetting Trim(gblApp_Name), "User", "User_Name", Trim(gblApp_Name)
            SaveSetting Trim(gblApp_Name), "User", "User_Password", "1007"
            SaveSetting Trim(gblApp_Name), "User", "User_Number", 0
            SaveSetting Trim(gblApp_Name), "User", "Last_User", 0
            UserRecord.User_Number = 0
            UserRecord.Name = "StartUp"
            UserRecord.Password = ""
            UserRecord.LastUser = 0
            cmbUsers.AddItem UserRecord.User_Number & " - " & (UserRecord.Name)
        Case Else
            cmbUsers.Clear
            'If user do exist look for other users in the database
            UserRecord.User_Number = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Number", Default:="StartUp"))
            UserRecord.Name = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Name", Default:="StartUp"))
            UserRecord.Password = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Password", Default:="StartUp"))
            UserRecord.LastUser = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="Last_User", Default:="0"))
            'Add user to splash List box
            cmbUsers.AddItem UserRecord.User_Number & " - " & (UserRecord.Name)
            If Trim(Server.SQL_Name) <> "" Then
                ActiveReadServer "Select user_No,User_Name from Users order by user_no"
                While Not rs.EOF
                    If rs.Fields("User_Name") = "New User" Then
                        SafetyCode "Users"
                    Else
                        cmbUsers.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    End If
                    rs.MoveNext
                Wend
                rs.Close
                ActiveReadServer "Select * from Branch_Details"
                If rs.RecordCount > 0 Then
                    PayoutPrint = Val(rs.Fields("PayoutPrint") & "")
                    ChargePrint = Val(rs.Fields("ChargePrint") & "")
                    RAPrint = Val(rs.Fields("RAPrint") & "")
                    Branch_Name = rs.Fields("Branch_Name")
                    Branch_No = rs.Fields("Branch_No")
                    Branch_Address = rs.Fields("Address")
                    Branch_Type = rs.Fields("Branch_Type")
                    Dept_Order = rs.Fields("Dept_Order")
                    Kitchen_Con = Abs(rs.Fields("Kitchen_Con"))
                    Conversion_description = rs.Fields("Conversion_Description")
                    If rs.Fields("Conversion_rate") <> 0 Then
                        Conversion_Rate = 1 / rs.Fields("Conversion_rate")
                    Else
                        Convertion_rate = 1
                    End If
                    Select Case rs.Fields("System_Access")
                        Case 1: System_Access = 1
                        Case 0: System_Access = 1
                        Case Else: System_Access = 1
                    End Select
                    Select Case rs.Fields("QCash")
                        Case 1: QCash = 1
                        Case 0: QCash = 0
                        Case Else: QCash = 0
                    End Select
                    Select Case rs.Fields("ChargeSlip")
                        Case 1:  ChargeSlip = 1
                        Case 0: ChargeSlip = 0
                        Case Else: ChargeSlip = 0
                    End Select
                    Select Case rs.Fields("System_Service")
                        Case 1: System_Service = 1
                        Case 0: System_Service = 0
                        Case Else: System_Service = 0
                    End Select
                    Vat_No = rs.Fields("Vat_No") & ""
                    Logo_File = rs.Fields("Logo_File") & ""
                    Swiss_Round = Val(rs.Fields("Swiss_Round") & "")
                    VoidReasons = Val(rs.Fields("Void_Reasons") & "")
                    StockBarcode = Val(rs.Fields("Stock_Barcode") & "")
                    TradePrint = Val(rs.Fields("Trade_Print") & "")
                    Member_No = Val(rs.Fields("Member_No") & "")
                    Zero_Print = Val(rs.Fields("Zero_Print") & "")
                    HappyHour = Val(rs.Fields("Happy_Active") & "")
                    If HappyHour = 1 Then
                        ActiveReadServer2 "Select Selling_Price from Happy_Hour"
                        If rs2.RecordCount > 0 Then
                            HappyHourPrice = rs2.Fields("Selling_Price")
                        End If
                        rs2.Close
                    End If
                    HappyHour1 = Val(rs.Fields("Happy_Active1") & "")
                    If HappyHour1 = 1 Then
                        ActiveReadServer2 "Select Selling_Price from Happy_Hour1"
                        If rs2.RecordCount > 0 Then
                            HappyHourPrice1 = rs2.Fields("Selling_Price")
                        End If
                        rs2.Close
                    End If
                End If
                rs.Close
                ActiveReadServer "Select * from Cost_Code"
                If rs.RecordCount > 0 Then
                    Cost_Code.One = rs.Fields("One") & ""
                    Cost_Code.Two = rs.Fields("Two") & ""
                    Cost_Code.Three = rs.Fields("Three") & ""
                    Cost_Code.Four = rs.Fields("Four") & ""
                    Cost_Code.Five = rs.Fields("Five") & ""
                    Cost_Code.Six = rs.Fields("Six") & ""
                    Cost_Code.Seven = rs.Fields("Seven") & ""
                    Cost_Code.Eight = rs.Fields("Eight") & ""
                    Cost_Code.Nine = rs.Fields("Nine") & ""
                    Cost_Code.Ten = rs.Fields("Zero") & ""
                End If
                rs.Close
                ActiveReadServer "Select Host_Name() as Comp_Name"
                Comp_Name = rs.Fields("Comp_Name")
                rs.Close
            End If
    End Select
    
    LogFiles.MainLog = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="Logs", key:="Main_Log"))
    LogFiles.ErrorLog = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="Logs", key:="Error_Log"))
    If Trim(UserRecord.LastUser) <> "" Then
        For i = 0 To cmbUsers.ListCount - 1
            If UserRecord.LastUser = Val(Left(cmbUsers.List(i), InStr(cmbUsers.List(i), "-") - 2)) Then
                cmbUsers.Text = cmbUsers.List(i)
                fmUser.Caption = cmbUsers.Text
                Exit For
            End If
        Next i
    End If
    Screen.MousePointer = 0
    If System_Access = 0 Then
        cmbUsers.Locked = True
    End If
    ActiveReadServer "Select * from Notice_Board order by Line_No"
    grdMess.Row = 0
    While Not rs.EOF
        grdMess.Row = grdMess.Row + 1
        grdMess.TextMatrix(grdMess.Row, 1) = rs.Fields("Description") & ""
        Select Case rs.Fields("Style") & ""
            Case "Sub Header"
                grdMess.TextMatrix(grdMess.Row, 0) = ">"
                grdMess.TextMatrix(grdMess.Row, 0) = ">"
                grdMess.Select grdMess.Row, 0, grdMess.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
            Case "Sub Header Underline"
                grdMess.TextMatrix(grdMess.Row, 0) = ">"
                grdMess.TextMatrix(grdMess.Row, 0) = ">"
                grdMess.Select grdMess.Row, 0, grdMess.Row, 1
                grdMess.CellBorder vbBlack, 0, 0, 0, 1, 0, 1
            Case "Normal"
                grdMess.TextMatrix(grdMess.Row, 0) = ""
                grdMess.TextMatrix(grdMess.Row, 0) = ""
                grdMess.Select grdMess.Row, 0, grdMess.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
            Case "Normal Underline"
                grdMess.TextMatrix(grdMess.Row, 0) = ""
                grdMess.TextMatrix(grdMess.Row, 0) = ""
                grdMess.Select grdMess.Row, 0, grdMess.Row, 1
                grdMess.CellBorder vbBlack, 0, 0, 0, 1, 0, 1
        End Select
        rs.MoveNext
    Wend
    rs.Close
    If Reservationsenabled = False Then cmdInput(18).Caption = ""
    If Reservationsenabled = True Then cmdInput(18).Caption = "Reservations"
    If Minitradeanalysis = True Then cmdInput(18).Caption = "Minitrade-analysis"
    If Newidea = True Then cmdInput(15).Caption = "New idea" ' used to be "Stock Requests"
    If Newidea = False Then cmdInput(15).Caption = "" ' "Stock Requests"
    
    
    If Devices.DisplayModel <> "" Then
        filenum = FreeFile
        Open Devices.DisplayPort For Output As filenum
        Print #filenum, Chr(27) & "@"
        Print #filenum, Chr(27) & " "; Format(Date, "yyyy-mm-dd") & "  " & Format(Time, "HH:MM AM/PM")
        Print #filenum, Chr(27) & "[B";
        Print #filenum, Chr(27) & "QD " & "HeroPOS"
        Close #filenum
    End If
    SafetyCode "Specials"
    On Error GoTo 0
End Sub



Private Sub Form_Load()
    On Error Resume Next
    

    'picAm.Visible = True
    lblVersion.Caption = "Version: " & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
    lblDate.Caption = Format(Date, "YYYY-MM-DD")
    'Look for default user
    Select Case GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Name", Default:="StartUp")
        Case "StartUp"
            'If user do not exists add a default user
            SaveSetting Trim(gblApp_Name), "User", "User_Name", Trim(gblApp_Name)
            SaveSetting Trim(gblApp_Name), "User", "User_Password", "1007"
            SaveSetting Trim(gblApp_Name), "User", "User_Number", 0
            SaveSetting Trim(gblApp_Name), "User", "Last_User", 0
            SaveSetting Trim(gblApp_Name), "Workstation", "Number", 0
            UserRecord.User_Number = 0
            UserRecord.Name = "StartUp"
            UserRecord.Password = ""
            UserRecord.LastUser = 0
            cmbUsers.AddItem UserRecord.User_Number & " - " & (UserRecord.Name)
        Case Else
            cmbUsers.Clear
            'If user do exist look for other users in the database
            UserRecord.User_Number = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Number", Default:="StartUp"))
            UserRecord.Name = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Name", Default:="StartUp"))
            UserRecord.Password = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="User_Password", Default:="StartUp"))
            UserRecord.LastUser = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="User", key:="Last_User", Default:="0"))
            Workstation_No = Val(Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Number", Default:="0")))
            'Add user to splash List box
            cmbUsers.AddItem UserRecord.User_Number & " - " & (UserRecord.Name)
            If Trim(Server.SQL_Name) <> "" Then
                ActiveReadServer "Select Workstation_Name from Workstations where Workstation_No= " & Workstation_No
                If rs.RecordCount > 0 Then
                    Workstation_Name = rs.Fields("Workstation_Name")
                Else
                    Workstation_Name = ""
                End If
                rs.Close
            End If
            If Trim(Server.SQL_Name) <> "" Then
                ActiveReadServer "Select user_No,User_Name from Users order by user_no"
                While Not rs.EOF
                    If rs.Fields("User_Name") = "New User" Then
                        SafetyCode "Users"
                    Else
                        cmbUsers.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                    End If
                    rs.MoveNext
                Wend
                rs.Close
                grdMess.Row = 0
                grdMess.Col = 0
                grdMess.ColWidth(0) = 180
                grdMess.TextMatrix(0, 0) = "Daily Notice Board"
                grdMess.TextMatrix(0, 1) = "Daily Notice Board"
                grdMess.MergeRow(0) = True
                grdMess.CellAlignment = flexAlignCenterCenter
                grdMess.Select 0, 0, 0, 1
                grdMess.CellBorder vbBlack, 0, 0, 0, 1, 0, 1
                grdMess.CellFontBold = True
            End If
    End Select
    
    LogFiles.MainLog = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="Logs", key:="Main_Log"))
    LogFiles.ErrorLog = Trim(GetSetting(appname:=Trim(gblApp_Name), Section:="Logs", key:="Error_Log"))
    If Trim(UserRecord.LastUser) <> "" Then
        For i = 0 To cmbUsers.ListCount - 1
            If UserRecord.LastUser = Val(Left(cmbUsers.List(i), InStr(cmbUsers.List(i), "-") - 2)) Then
                cmbUsers.Text = cmbUsers.List(i)
                fmUser.Caption = cmbUsers.Text
                Exit For
            End If
        Next i
    End If
    If Trim(Server.SQL_Name) <> "" Then
        Load frmSlipDetails
    End If

    On Error GoTo 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdChange.Visible = False
End Sub



Private Sub Timer1_Timer()
Maindate = DateValue(Date)
On Error GoTo 0

If chkdateagain = "" Then
Dim datestring As Variant
Dim currentdatestring As Variant
If rs6.State = 1 Then rs6.Close
ActiveReadServer6 " select * from Subscription"
If rs6.RecordCount > 0 Then
chkdateagain = Format(rs6.Fields("Subscriptiondate"), "DD-MMMM-YYYY")
datestring = Format(chkdateagain, "yyyy/mm/dd")
currentdatestring = Format(Date, "yyyy/mm/dd")
lblenddate.Caption = chkdateagain
currentdatestring = DateValue(currentdatestring)
datestring = DateValue(datestring) - 31
If DateValue(currentdatestring) < DateValue(datestring) Then
lblenddate.ForeColor = vbWhite
End If
If DateValue(currentdatestring) > DateValue(datestring) Then
lblenddate.ForeColor = vbRed
lblenddate.Caption = chkdateagain + "  Warning " & DateValue(datestring + 31) - DateValue(currentdatestring) & "days to register!"
End If
End If
End If
End Sub



Private Sub Timer2_Timer()
chkdateagain = ""
End Sub

Private Sub Timer3_Timer()
    lblTime.Caption = Time
    lblDate.Caption = Format(Date, "DD MMM YYYY")
    
    If DateValue(Date) > DateValue(Maindate) Then chkdateagain = ""
    If DateValue(Date) < DateValue(Maindate) Then chkdateagain = ""
    Process_Running = False 'Reset running process flag - prevents dead lock
End Sub
Private Sub Timer4_Timer()
    Select Case cmdErr.BackColor
        Case &H30A8CF
            cmdErr.BackColor = &H40C0&
        Case &H40C0&
            cmdErr.BackColor = &H30A8CF
    End Select
End Sub
Private Sub Timer5_Timer()
    On Error GoTo trap
    If Tel_Ex <> 0 Then
        readTelEx
    End If
    If Branch_Type < 8 Then
        ActiveReadServer4 "Select Happy_Active from Branch_Details"
        If rs4.RecordCount > 0 Then
            HappyHour = Val(rs4.Fields("Happy_Active") & "")
        End If
        rs4.Close
        Select Case Format(Date, "DDD")
            Case "Mon"
                ActiveReadServer4 "Select * from Happy_Hour where Week_Day = 1 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active = 1"
                                HappyHourPrice = rs4.Fields("Selling_Price")
                                HappyHour = 1
                            Else
                                If HappyHour = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active = 0"
                                    HappyHour = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Tue"
                ActiveReadServer4 "Select * from Happy_Hour where Week_Day = 2 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active = 1"
                                HappyHourPrice = rs4.Fields("Selling_Price")
                                HappyHour = 1
                            Else
                                If HappyHour = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active = 0"
                                    HappyHour = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Wed"
                ActiveReadServer4 "Select * from Happy_Hour where Week_Day = 3 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active = 1"
                                HappyHourPrice = rs4.Fields("Selling_Price")
                                HappyHour = 1
                            Else
                                If HappyHour = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active = 0"
                                    HappyHour = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Thu"
                ActiveReadServer4 "Select * from Happy_Hour where Week_Day = 4 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active = 1"
                                HappyHourPrice = rs4.Fields("Selling_Price")
                                HappyHour = 1
                            Else
                                If HappyHour = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active = 0"
                                    HappyHour = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Fri"
                ActiveReadServer4 "Select * from Happy_Hour where Week_Day = 5 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active = 1"
                                HappyHourPrice = rs4.Fields("Selling_Price")
                                HappyHour = 1
                            Else
                                If HappyHour = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active = 0"
                                    HappyHour = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Sat"
                ActiveReadServer4 "Select * from Happy_Hour where Week_Day = 6 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active = 1"
                                HappyHourPrice = rs4.Fields("Selling_Price")
                                HappyHour = 1
                            Else
                                If HappyHour = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active = 0"
                                    HappyHour = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Sun"
                ActiveReadServer4 "Select * from Happy_Hour where Week_Day = 7 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active = 1"
                                HappyHourPrice = rs4.Fields("Selling_Price")
                                HappyHour = 1
                            Else
                                If HappyHour = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active = 0"
                                    HappyHour = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
        End Select
        
        ActiveReadServer4 "Select Happy_Active1 from Branch_Details"
        If rs4.RecordCount > 0 Then
            HappyHour1 = Val(rs4.Fields("Happy_Active1") & "")
        End If
        rs4.Close
        Select Case Format(Date, "DDD")
            Case "Mon"
                ActiveReadServer4 "Select * from Happy_Hour1 where Week_Day = 1 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 1"
                                HappyHourPrice1 = rs4.Fields("Selling_Price")
                                HappyHour1 = 1
                            Else
                                If HappyHour1 = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 0"
                                    HappyHour1 = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Tue"
                ActiveReadServer4 "Select * from Happy_Hour1 where Week_Day = 2 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 1"
                                HappyHourPrice1 = rs4.Fields("Selling_Price")
                                HappyHour1 = 1
                            Else
                                If HappyHour1 = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 0"
                                    HappyHour1 = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Wed"
                ActiveReadServer4 "Select * from Happy_Hour1 where Week_Day = 3 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 1"
                                HappyHourPrice1 = rs4.Fields("Selling_Price")
                                HappyHour1 = 1
                            Else
                                If HappyHour1 = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 0"
                                    HappyHour1 = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Thu"
                ActiveReadServer4 "Select * from Happy_Hour1 where Week_Day = 4 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 1"
                                HappyHourPrice1 = rs4.Fields("Selling_Price")
                                HappyHour1 = 1
                            Else
                                If HappyHour1 = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 0"
                                    HappyHour1 = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Fri"
                ActiveReadServer4 "Select * from Happy_Hour1 where Week_Day = 5 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 1"
                                HappyHourPrice1 = rs4.Fields("Selling_Price")
                                HappyHour1 = 1
                            Else
                                If HappyHour1 = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 0"
                                    HappyHour1 = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Sat"
                ActiveReadServer4 "Select * from Happy_Hour1 where Week_Day = 6 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 1"
                                HappyHourPrice1 = rs4.Fields("Selling_Price")
                                HappyHour1 = 1
                            Else
                                If HappyHour1 = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 0"
                                    HappyHour1 = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
            Case "Sun"
                ActiveReadServer4 "Select * from Happy_Hour1 where Week_Day = 7 and Auto_Change = 1"
                If rs4.RecordCount > 0 Then
                    If Val(rs4.Fields("Active") & "") = 1 Then
                        If Format(Time, "HH:MM:SS") > Format(rs4.Fields("Start_Time"), "HH:MM:SS") Then
                            If Format(Time, "HH:MM:SS") < Format(rs4.Fields("Stop_Time"), "HH:MM:SS") Then
                                ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 1"
                                HappyHourPrice1 = rs4.Fields("Selling_Price")
                                HappyHour1 = 1
                            Else
                                If HappyHour1 = 1 Then
                                    ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 0"
                                    HappyHour1 = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs4.Close
        End Select
    End If
trap:
    On Error GoTo 0
End Sub
Private Sub readTelEx()
    Dim AllColumns As Variant
    On Error GoTo trap
    filenum = FreeFile
    Open Tel_Ex_Dir For Input As filenum
    Do While Not EOF(filenum)   ' Loop until end of file.
        DoEvents
        Line Input #filenum, TextLine   ' Read line into variable.
        AllColumns = Split(TextLine, ",")
        If AllColumns(2) <> "" Then
            Room_No = Right(Trim(AllColumns(2)), 2)
            Tel_No = Trim(AllColumns(1))
            Duration = Trim(AllColumns(3))
            TimeofCall = Trim(AllColumns(8))
            Price = Val(Trim(Mid(AllColumns(10), 1, InStr(AllColumns(10), ";") - 1)))
            Doc_No = Trim(AllColumns(6))
            Res_No = 0
            ActiveReadServer "Select * from Tel_Ex where Room_No= " & Val(Room_No)
            If rs.RecordCount > 0 Then
                ActiveReadServer1 "Select Res_No from Reservations where Res_Type = 2 and Room_No= " & Val(Room_No)
                If rs1.RecordCount > 0 Then
                    Res_No = Val(rs1.Fields("Res_no") & "")
                End If
                rs1.Close
                If Res_No <> 0 Then
                    ActiveReadServer "Select Balance from Room_Accounts where Res_No =" & Res_No
                    If rs.RecordCount > 0 Then
                        rs.MoveLast
                        OldBalance = Val(rs.Fields("Balance") & "")
                    Else
                        OldBalance = 0
                    End If
                    rs.Close
                    ActiveUpdateServer "INSERT INTO [Room_Accounts]([Transaction_Type],[Date_Time], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Res_No])" & _
                    "VALUES('Telephone - " & Tel_No & "',Getdate()," & Doc_No & ",'" & Val(Room_No) & "'," & Price & ",0," & OldBalance + Price & "," & Res_No & ")"
                End If
            End If
            rs.Close
        End If
    Loop
    Close #filenum   ' Close
    DoEvents
    Kill Tel_Ex_Dir
    On Error GoTo 0
    Exit Sub
trap:
    On Error GoTo 0
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer) 'Entry point
Savepassword = ""
If fmClock.Visible = True Then fmClock.Visible = False
If System_Access = 1 And KeyCode = 13 Then
    ActiveReadServer "Select User_No,User_Name from Users where User_Password = '" & txtPassword.Text & "'"
    If rs.RecordCount > 0 Then
        newText = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
        rs.Close
        Savepassword = txtPassword.Text
        cmbUsers.Text = newText
        fmUser.Caption = cmbUsers.Text
        DoEvents
        txtPassword.Text = Savepassword
    Else
        rs.Close
    End If
End If
Select Case KeyCode
        Case 13
            Select Case txtPassword.Tag
                Case "Stock Requests"
                    KeyCode = 0
                    Timer4.Enabled = False
                    cmdErr.BackColor = &H30A8CF
                    cmdErr.Visible = False
                    txtPassword.Tag = ""
                    If cmbUsers.Text = "" Then
                        Load frmError
                        frmError.lblCap.Caption = "Please select a user name before entering your password."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        cmbUsers.SetFocus
                        Exit Sub
                    End If
                    If Trim(UserRecord.Password) = Trim(txtPassword) Then
                        ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                        "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                        If rs.RecordCount > 0 Then
                            If rs.Fields("Function_Key") = 4 Then
                                Load frmError
                                frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Making Stock Requests"
                                frmError.Show vbModal
                                txtPassword.SetFocus
                                txtPassword.Text = ""
                                rs.Close
                                Exit Sub
                            End If
                        Else
                            Load frmError
                            frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Making Stock Requests"
                            frmError.Show vbModal
                            txtPassword.SetFocus
                            txtPassword.Text = ""
                            rs.Close
                            Exit Sub
                        End If
                        rs.Close
                        UserRecord.LastUser = Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2))
                        Select Case UserRecord.Inventory
                            Case False
                                Load frmError
                                frmError.lblCap.Caption = "You do not have access rights to make Stock Requests."
                                frmError.Show vbModal
                                txtPassword.Text = ""
                                txtPassword.SetFocus
                                Exit Sub
                            Case True
                                Screen.MousePointer = 11
                                KeyCode = 0
                                GlobalMode = TillMode.StartMode
                                frmRequest.Show vbModal
                                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Stock Requests')"
                                DoEvents
                                Screen.MousePointer = 0
                        End Select
                        DoEvents
                        SaveSetting Trim(gblApp_Name), "User", "Last_User", UserRecord.LastUser
                        If Trim(Server.SQL_Name) = "" Then
                            frmSettings.Show vbModal
                        End If
                    Else
                        Load frmError
                        frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                        Exit Sub
                    End If
                    Exit Sub
                Case "Reservations"
                    KeyCode = 0
                    Timer4.Enabled = False
                    cmdErr.BackColor = &H30A8CF
                    cmdErr.Visible = False
                    txtPassword.Tag = ""
                    If cmbUsers.Text = "" Then
                        Load frmError
                        frmError.lblCap.Caption = "Please select a user name before entering your password."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        cmbUsers.SetFocus
                        Exit Sub
                    End If
                    If Trim(UserRecord.Password) = Trim(txtPassword) Then
                        ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                        "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                        If rs.RecordCount > 0 Then
                            If rs.Fields("Function_Key") = 4 Then
                                Load frmError
                                frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Making Reservations"
                                frmError.Show vbModal
                                txtPassword.SetFocus
                                txtPassword.Text = ""
                                rs.Close
                                Exit Sub
                            End If
                        Else
                            Load frmError
                            frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Making Reservations"
                            frmError.Show vbModal
                            txtPassword.SetFocus
                            txtPassword.Text = ""
                            rs.Close
                            Exit Sub
                        End If
                        rs.Close
                        UserRecord.LastUser = Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2))
                        Select Case UserRecord.Reports
                            Case False
                                Load frmError
                                frmError.lblCap.Caption = "You do not have access rights to make Reservations."
                                frmError.Show vbModal
                                txtPassword.Text = ""
                                txtPassword.SetFocus
                                Exit Sub
                            Case True
                                Screen.MousePointer = 11
                                KeyCode = 0
                                GlobalMode = TillMode.StartMode
                                frmRestRes.Show
                                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Reservations')"
                                DoEvents
                                Me.Hide
                                Screen.MousePointer = 0
                        End Select
                        DoEvents
                        SaveSetting Trim(gblApp_Name), "User", "Last_User", UserRecord.LastUser
                        If Trim(Server.SQL_Name) = "" Then
                            frmSettings.Show vbModal
                        End If
                    Else
                        Load frmError
                        frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                        Exit Sub
                    End If
                    Exit Sub
                Case "Clock In"
                    Timer4.Enabled = False
                    cmdErr.BackColor = &H30A8CF
                    cmdErr.Visible = False
                    txtPassword.Tag = ""
                    If Trim(UserRecord.Password) = Trim(txtPassword) Then
                        ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                        "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                        If rs.RecordCount > 0 Then
                            If rs.Fields("Function_Key") = 3 Then
                                Load frmError
                                frmError.lblCap.Caption = "You have already Clocked In. Enter password and click OK or Clock Out."
                                frmError.Show vbModal
                                txtPassword.SetFocus
                                txtPassword.Text = ""
                                rs.Close
                                Exit Sub
                            End If
                        End If
                        rs.Close
                        If Devices.TwoDrawer = 1 And UserRecord.uType = 4 Then
                            frmSplash.cmbUsers.Tag = ""
                            Load frmChooseRes
                            frmSplash.Tag = "Not Now"
                            frmChooseRes.Caption = "Please Select a Drawer for this Shift."
                            DoEvents
                            frmChooseRes.Show vbModal
                            Select Case frmSplash.cmbUsers.Tag
                                Case "Drawer One"
                                    Drawer = 1
                                Case "Drawer Two"
                                    Drawer = 2
                                Case Else
                                    frmError.lblCap.Caption = "You have to Select a Drawer"
                                    frmError.Show vbModal
                                    txtPassword.SetFocus
                                    txtPassword.Text = ""
                                    Exit Sub
                            End Select
                        Else
                            Drawer = 1
                        End If
                        fmClock.Caption = UserRecord.Name & " Clocked in at " & Format(Time, "HH:MM")
                        fmClock.Visible = True
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),3," & Workstation_No & ")"
                        ActiveUpdateServer "Update Users set Workstation_No=" & Workstation_No & ",Drawer_No=" & Drawer & ",Clocked_In = 1 where User_No = " & UserRecord.User_Number
                        KeyCode = 0
                        DoEvents
                        Print_Clock_In
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                    Else
                        Load frmError
                        frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                        Exit Sub
                    End If
                    Exit Sub
                Case "Clock Out"
                    Timer4.Enabled = False
                    cmdErr.BackColor = &H30A8CF
                    cmdErr.Visible = False
                    txtPassword.Tag = ""
                    If Trim(UserRecord.Password) = Trim(txtPassword) Then
                        ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                        "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                        If rs.RecordCount > 0 Then
                            If rs.Fields("Function_Key") = 4 Then
                                Load frmError
                                frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Clocking Out"
                                frmError.Show vbModal
                                txtPassword.SetFocus
                                txtPassword.Text = ""
                                rs.Close
                                Exit Sub
                            End If
                        Else
                            Load frmError
                            frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Clocking Out"
                            frmError.Show vbModal
                            txtPassword.SetFocus
                            txtPassword.Text = ""
                            rs.Close
                            Exit Sub
                        End If
                        rs.Close
                        ActiveReadServer "Select * from Cash_ups where Finalized = 0 and User_No = " & UserRecord.User_Number
                        If rs.RecordCount > 0 Then
                            rs.Close
                            Load frmError
                            frmError.lblCap.Caption = "You cannot Clock Out while you have an Outstanding Cashup."
                            frmError.Show vbModal
                            txtPassword.Text = ""
                            txtPassword.SetFocus
                            Exit Sub
                        End If
                        rs.Close
                        
                        ActiveReadServer "Select count(Table_No) as TableCount from Table_Listing_View where User_No = " & UserRecord.User_Number
                        Select Case rs.Fields("TableCount")
                            Case 0
                            Case 1
                                Load frmError
                                frmError.lblCap.Caption = "You still have an Open Table. You cannot Clock Out while you have an Open Table."
                                frmError.Show vbModal
                                txtPassword.Text = ""
                                txtPassword.SetFocus
                                rs.Close
                                Exit Sub
                            Case Else
                                Load frmError
                                cmdErr.Caption = "You still have " & rs.Fields("TableCount") & " Open Tables. You cannot Clock Out while you have an Open Table."""
                                frmError.Show vbModal
                                txtPassword.Text = ""
                                txtPassword.SetFocus
                                rs.Close
                                Exit Sub
                        End Select
                        rs.Close
                        
                        ActiveReadServer "Select count(Table_No) as TableCount from Table_Listing_View where User_No <> " & Mid(cmbUsers.Text, 1, InStr(cmbUsers.Text, "-") - 1) & _
                        " and Previous_Owner = " & Mid(cmbUsers.Text, 1, InStr(cmbUsers.Text, "-") - 1)
                        If rs.Fields("TableCount") > 0 Then
                            Timer2.Enabled = True
                            Select Case rs.Fields("TableCount")
                                Case 1
                                    cmdErr.Caption = "A Transfer by this User has nor been accepted."
                                Case Else
                                    cmdErr.Caption = rs.Fields("TableCount") & " Transfers by this User has nor been accepted."
                            End Select
                            cmdErr.Visible = True
                            rs.Close
                            Exit Sub
                        End If
                        
                        fmClock.Caption = UserRecord.Name & " Clocked Out at " & Format(Time, "HH:MM")
                        Print_Clock_Out
                        fmClock.Visible = True
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),4," & Workstation_No & ")"
                        ActiveUpdateServer "Update Users set Drawer_No = 0, Clocked_In = 0 where User_No = " & UserRecord.User_Number
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                    Else
                        Load frmError
                        frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                    End If
                    Exit Sub
                Case "Cash-up's"
                    KeyCode = 0
                    Timer4.Enabled = False
                    cmdErr.BackColor = &H30A8CF
                    cmdErr.Visible = False
                    txtPassword.Tag = ""
                    If cmbUsers.Text = "" Then
                        Load frmError
                        frmError.lblCap.Caption = "Please select a user name before entering your password."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        cmbUsers.SetFocus
                        Exit Sub
                    End If
                    If Trim(UserRecord.Password) = Trim(txtPassword) Then
                        UserRecord.LastUser = Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2))
                        Select Case UserRecord.Reports
                            Case False
                                ActiveReadServer "Select Cashup_No from Counters where Finalized= 0 and User_No=" & UserRecord.User_Number
                                If rs.RecordCount > 0 Then
                                    Cashup_No = rs.Fields("Cashup_No")
                                    rs.Close
                                    
                                    ActiveReadServer "Select count(Table_No) as TableCount from Table_Listing_View where User_No = " & UserRecord.User_Number
                                    If rs.Fields("TableCount") > 0 Then
                                        Timer4.Enabled = True
                                        Select Case rs.Fields("TableCount")
                                            Case 1
                                                cmdErr.Caption = "You still have an Open Table."
                                            Case Else
                                                cmdErr.Caption = "You still have " & rs.Fields("TableCount") & " Open Tables."
                                        End Select
                                        cmdErr.Visible = True
                                        txtPassword.Text = ""
                                        rs.Close
                                        Exit Sub
                                    End If
                                    rs.Close
                                    
                                    ActiveReadServer "Select count(Tab_No) as TabCount from Tab_Listing_View where User_No = " & UserRecord.User_Number
                                    If rs.Fields("TabCount") > 0 Then
                                        Timer4.Enabled = True
                                        Select Case rs.Fields("TabCount")
                                            Case 1
                                                cmdErr.Caption = "You still have an Open Tab."
                                            Case Else
                                                cmdErr.Caption = "You still have " & rs.Fields("TabCount") & " Open Tabs."
                                        End Select
                                        cmdErr.Visible = True
                                        txtPassword.Text = ""
                                        rs.Close
                                        Exit Sub
                                    End If
                                    rs.Close
                                    
                                    Screen.MousePointer = 11
                                    Load frmCapture
                                    frmCapture.Tag = -10
                                    frmCapture.lblHeading.Tag = Cashup_No
                                    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Cash-up Capture')"
                                    Screen.MousePointer = 0
                                    frmCapture.Show vbModal
                                    frmCapture.lblHeading.Tag = ""
                                    Exit Sub
                                Else
                                    rs.Close
                                    Load frmError
                                    frmError.lblCap.Caption = "You do not have access rights to use Cash-up's."
                                    frmError.Show vbModal
                                    txtPassword.Text = ""
                                    txtPassword.SetFocus
                                    Exit Sub
                                End If
                            Case True
                                Screen.MousePointer = 11
                                KeyCode = 0
                                GlobalMode = TillMode.CashupMode
                                frmTillReport.Show
                                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Cash-up`s')"
                                DoEvents
                                Me.Hide
                                Screen.MousePointer = 0
                        End Select
                        DoEvents
                        SaveSetting Trim(gblApp_Name), "User", "Last_User", UserRecord.LastUser
                        If Trim(Server.SQL_Name) = "" Then
                            frmSettings.Show vbModal
                        End If
                    Else
                        Load frmError
                        frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                        frmError.Show vbModal
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                        Exit Sub
                    End If
                    Exit Sub
            End Select
            KeyCode = 0
            If cmbUsers.Text = "" Then
                Load frmError
                frmError.lblCap.Caption = "Please select a user name before entering your password."
                frmError.Show vbModal
                txtPassword.Text = ""
                On Error Resume Next
                cmbUsers.SetFocus
                On Error GoTo 0
                Exit Sub
            End If
            If txtPassword = "000000" Then
                txtPassword.Text = ""
                Exit Sub
            End If
            If Trim(UserRecord.Password) = Trim(txtPassword.Text) Then
                ActiveReadServer "Select Function_Key from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                If rs.RecordCount > 0 Then
                    If rs.Fields("Function_Key") = 4 Then
                        Load frmError
                        frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Logging On"
                        frmError.Show vbModal
                        txtPassword.SetFocus
                        txtPassword.Text = ""
                        rs.Close
                        Exit Sub
                    End If
                Else
                    Load frmError
                    frmError.lblCap.Caption = "You have not Clocked In. Please Clock In first before Logging On"
                    frmError.Show vbModal
                    txtPassword.SetFocus
                    txtPassword.Text = ""
                    rs.Close
                    Exit Sub
                End If
                rs.Close
                UserRecord.LastUser = Val(Left(cmbUsers.Text, InStr(cmbUsers.Text, "-") - 2))
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),1," & Workstation_No & ")"
                Clear_TillData  'Kotie 17-03-2013  21:00
                Select Case UserRecord.uType
                    Case 2
                        
                        checked = Checksubscriptiondb
                        If checked = True Then
                        If frmSplash.Height < 10000 Then
                            Screen.MousePointer = 11
                            KeyCode = 0
                            GlobalMode = TillMode.StartMode
                            Load frmBar
                            frmBar.picSlip.Visible = False
                            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Sales')"
                            DoEvents
                            Finalizing = False
                            frmBar.Show
                            Screen.MousePointer = 0
                            Else
                        MsgBox " Your subscription to HeroPOS has expired please inform HeroPOS to recieve a new date key"
                        End If
                        Else
                            Screen.MousePointer = 11
                            KeyCode = 0
                            Load frmMain
                            For i = 0 To frmMain.cmdBar.Count - 1
                                frmMain.cmdBar(i).Enabled = True
                            Next i
                            WriteMainLog
                            DoEvents
                            For Each Form In Forms
                                Select Case Form.Name
                                    Case "frmUsers", "frmGuest", "frmRooms", "frmCheck", "frmReports", "frmRes"
                                        Unload Form
                                End Select
                            Next
                            frmDetails.Show
                            frmSplash.Hide
                            Screen.MousePointer = 0
                        End If
                    Case 8
                        checked = Checksubscriptiondb
                        If checked = True Then
                        Me.Hide
                        KeyCode = 0
                        DoEvents
                        GlobalMode = TillMode.StartMode
                        Finalizing = False
                        frmSales.Show
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Sales')"
                        DoEvents
                        Else
                        MsgBox " Your subscription to HeroPOS has expired please inform HeroPOS to recieve a new date key"
                        End If
                    Case 4
                        Checksubscriptiondb
                        checked = Checksubscriptiondb
                        If checked = True Then
                        Screen.MousePointer = 11
                        KeyCode = 0
                        GlobalMode = TillMode.StartMode
                        Load frmBar
                        frmBar.picSlip.Visible = False
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Sales')"
                        DoEvents
                        Finalizing = False
                        frmBar.Show
                        Screen.MousePointer = 0
                       Else
                        MsgBox " Your subscription to HeroPOS has expired please inform HeroPOS to recieve a new date key"
                        End If
                    Case 3
                        checked = Checksubscriptiondb
                        If checked = True Then
                        Screen.MousePointer = 11
                        KeyCode = 0
                        GlobalMode = TillMode.StartMode
                        Load frmSales1
                        ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Sales')"
                        DoEvents
                        frmInput.Show
                        Screen.MousePointer = 0
                        Me.Hide
                        Else
                        MsgBox " Your subscription to HeroPOS has expired please inform HeroPOS to recieve a new date key"
                        End If
                    Case 10
                        Load frmError
                        frmError.lblCap.Caption = "Staff Members cannot Log on. They can only Clock In and Out for Shift Clocking"
                        frmError.Show vbModal
                        txtPassword.SetFocus
                        txtPassword.Text = ""
                        Exit Sub
                    Case Else
                        Screen.MousePointer = 11
                        KeyCode = 0
                        Load frmMain
                        For i = 0 To frmMain.cmdBar.Count - 1
                            frmMain.cmdBar(i).Enabled = True
                        Next i
                        WriteMainLog
                        DoEvents
                        For Each Form In Forms
                            Select Case Form.Name
                                Case "frmUsers", "frmGuest", "frmRooms", "frmCheck", "frmReports", "frmRes"
                                    Unload Form
                            End Select
                        Next
                        frmDetails.Show
                        frmSplash.Hide
                        Screen.MousePointer = 0
                End Select
                DoEvents
                SaveSetting Trim(gblApp_Name), "User", "Last_User", UserRecord.LastUser
                If Trim(Server.SQL_Name) = "" Then
                    frmSettings.Show vbModal
                End If
            Else
                Load frmError
                frmError.lblCap.Caption = "You have entered an incorrect password. Please retry."
                frmError.Show vbModal
                txtPassword.Text = ""
                txtPassword.SetFocus
                Exit Sub
            End If
        Case 27
            'End
    End Select
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 27, 48 To 57
    Case Else
        KeyAscii = 0
End Select
End Sub
Private Sub Print_Clock_Out()
If Slip_Printer = "<None>" Then GoTo Done1
On Error GoTo trap
With frmSlipDetails
    PrintErr = 0
    Slip_Port = ""

    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As #filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
        
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
    End If

    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "USER CLOCK OUT"
    
    Print #filenum, Trim(UserRecord.User_Number & " - " & UserRecord.Name)
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    Print #filenum, "Date: " & Format(Date, "YYYY-MM-DD DDDD")
    Print #filenum, "Time: " & Time
    Print #filenum, String(33, "-")
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, "PRESENTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "ACCEPTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);

    For i = 1 To .grdFoot.Rows - 1
        If Trim(.grdFoot.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdFoot.TextMatrix(i, 1))
                Case "Line Feeds"
                    Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdFoot.TextMatrix(i, 0)));
                Case Else
                    Select Case .grdFoot.TextMatrix(i, 2)
                        Case "Left": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdFoot.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #filenum
End With
On Error GoTo 0
Done1:
Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim x As Printer
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    frmError.Show vbModal
    Close #filenum
    On Error GoTo 0
End Sub
Private Sub Print_Clock_In()
If Slip_Printer = "<None>" Then GoTo Done2
On Error GoTo trap
With frmSlipDetails
    PrintErr = 0
    Slip_Port = ""

    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then  ' Kotie 17-03 2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As #filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
    End If

    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "USER CLOCK IN"
    
    Print #filenum, Trim(UserRecord.User_Number & " - " & UserRecord.Name)
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    Print #filenum, "Date: " & Format(Date, "YYYY-MM-DD DDDD")
    Print #filenum, "Time: " & Time
    Print #filenum, String(33, "-")
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, "PRESENTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "ACCEPTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);

    For i = 1 To .grdFoot.Rows - 1
        If Trim(.grdFoot.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdFoot.TextMatrix(i, 1))
                Case "Line Feeds"
                    Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdFoot.TextMatrix(i, 0)));
                Case Else
                    Select Case .grdFoot.TextMatrix(i, 2)
                        Case "Left": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdFoot.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #filenum
End With
On Error GoTo 0
Done2:
Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim x As Printer
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    frmError.Show vbModal
    Close #filenum
    On Error GoTo 0
End Sub
