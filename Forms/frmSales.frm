VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSales 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11415
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H002298CA&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSales.frx":0000
   ScaleHeight     =   11415
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2190
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picSearch 
      BackColor       =   &H0071B9DB&
      Height          =   1305
      Left            =   330
      ScaleHeight     =   1245
      ScaleWidth      =   14745
      TabIndex        =   40
      Top             =   330
      Visible         =   0   'False
      Width           =   14805
      Begin BTNENHLib4.BtnEnh cmdArr 
         Height          =   1080
         Index           =   0
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   945
         _Version        =   524298
         _ExtentX        =   1667
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "3"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   30
            Charset         =   2
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
         CornerFactor    =   12
         Surface         =   1
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":E36D
         textLT          =   "frmSales.frx":E3CF
         textCT          =   "frmSales.frx":E3E7
         textRT          =   "frmSales.frx":E3FF
         textLM          =   "frmSales.frx":E417
         textRM          =   "frmSales.frx":E42F
         textLB          =   "frmSales.frx":E447
         textCB          =   "frmSales.frx":E45F
         textRB          =   "frmSales.frx":E477
         colorBack       =   "frmSales.frx":E48F
         colorIntern     =   "frmSales.frx":E4B9
         colorMO         =   "frmSales.frx":E4E3
         colorFocus      =   "frmSales.frx":E50D
         colorDisabled   =   "frmSales.frx":E537
         colorPressed    =   "frmSales.frx":E561
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin VB.PictureBox picKey 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   13650
         ScaleHeight     =   1065
         ScaleWidth      =   1095
         TabIndex        =   43
         Top             =   1980
         Width           =   1095
         Begin MSForms.Image cmdKeyboard 
            Height          =   945
            Left            =   100
            ToolTipText     =   " <Ctrl-K> "
            Top             =   70
            Width           =   945
            BorderColor     =   16777215
            Size            =   "1667;1667"
         End
         Begin MSForms.Image Image2 
            Height          =   1065
            Left            =   30
            Top             =   0
            Width           =   1095
            BackColor       =   2267338
            BorderStyle     =   0
            SpecialEffect   =   1
            Size            =   "1931;1879"
         End
      End
      Begin btButtonEx.ButtonEx cmdArrow1 
         Height          =   795
         Index           =   1
         Left            =   13710
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   10080
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1402
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
      Begin btButtonEx.ButtonEx cmdArrow1 
         Height          =   765
         Index           =   0
         Left            =   13710
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1349
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
      Begin BTNENHLib4.BtnEnh cmdArr 
         Height          =   1080
         Index           =   1
         Left            =   10710
         TabIndex        =   44
         Top             =   0
         Width           =   945
         _Version        =   524298
         _ExtentX        =   1667
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "4"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   30
            Charset         =   2
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
         CornerFactor    =   12
         Surface         =   1
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":E58B
         textLT          =   "frmSales.frx":E5ED
         textCT          =   "frmSales.frx":E605
         textRT          =   "frmSales.frx":E61D
         textLM          =   "frmSales.frx":E635
         textRM          =   "frmSales.frx":E64D
         textLB          =   "frmSales.frx":E665
         textCB          =   "frmSales.frx":E67D
         textRB          =   "frmSales.frx":E695
         colorBack       =   "frmSales.frx":E6AD
         colorIntern     =   "frmSales.frx":E6D7
         colorMO         =   "frmSales.frx":E701
         colorFocus      =   "frmSales.frx":E72B
         colorDisabled   =   "frmSales.frx":E755
         colorPressed    =   "frmSales.frx":E77F
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdClose 
         Height          =   1080
         Left            =   13590
         TabIndex        =   45
         Top             =   0
         Width           =   1155
         _Version        =   524298
         _ExtentX        =   2037
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "X"
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
         CornerFactor    =   12
         BackColorContainer=   12632256
         SpecialEffect   =   1
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":E7A9
         textLT          =   "frmSales.frx":E80B
         textCT          =   "frmSales.frx":E823
         textRT          =   "frmSales.frx":E83B
         textLM          =   "frmSales.frx":E853
         textRM          =   "frmSales.frx":E86B
         textLB          =   "frmSales.frx":E883
         textCB          =   "frmSales.frx":E89B
         textRB          =   "frmSales.frx":E8B3
         colorBack       =   "frmSales.frx":E8CB
         colorIntern     =   "frmSales.frx":E8F5
         colorMO         =   "frmSales.frx":E91F
         colorFocus      =   "frmSales.frx":E949
         colorDisabled   =   "frmSales.frx":E973
         colorPressed    =   "frmSales.frx":E99D
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdDeptStrip 
         Height          =   1080
         Index           =   0
         Left            =   930
         TabIndex        =   48
         Top             =   0
         Width           =   1425
         _Version        =   524298
         _ExtentX        =   2514
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "1"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":E9C7
         textLT          =   "frmSales.frx":EA29
         textCT          =   "frmSales.frx":EA41
         textRT          =   "frmSales.frx":EA59
         textLM          =   "frmSales.frx":EA71
         textRM          =   "frmSales.frx":EA89
         textLB          =   "frmSales.frx":EAA1
         textCB          =   "frmSales.frx":EAB9
         textRB          =   "frmSales.frx":EAD1
         colorBack       =   "frmSales.frx":EAE9
         colorIntern     =   "frmSales.frx":EB13
         colorMO         =   "frmSales.frx":EB3D
         colorFocus      =   "frmSales.frx":EB67
         colorDisabled   =   "frmSales.frx":EB91
         colorPressed    =   "frmSales.frx":EBBB
         Style           =   2
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdDeptStrip 
         Height          =   1080
         Index           =   5
         Left            =   7920
         TabIndex        =   49
         Top             =   0
         Width           =   1395
         _Version        =   524298
         _ExtentX        =   2461
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "1"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":EBE5
         textLT          =   "frmSales.frx":EC47
         textCT          =   "frmSales.frx":EC5F
         textRT          =   "frmSales.frx":EC77
         textLM          =   "frmSales.frx":EC8F
         textRM          =   "frmSales.frx":ECA7
         textLB          =   "frmSales.frx":ECBF
         textCB          =   "frmSales.frx":ECD7
         textRB          =   "frmSales.frx":ECEF
         colorBack       =   "frmSales.frx":ED07
         colorIntern     =   "frmSales.frx":ED31
         colorMO         =   "frmSales.frx":ED5B
         colorFocus      =   "frmSales.frx":ED85
         colorDisabled   =   "frmSales.frx":EDAF
         colorPressed    =   "frmSales.frx":EDD9
         Style           =   2
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdDeptStrip 
         Height          =   1080
         Index           =   3
         Left            =   5130
         TabIndex        =   50
         Top             =   0
         Width           =   1395
         _Version        =   524298
         _ExtentX        =   2461
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "1"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":EE03
         textLT          =   "frmSales.frx":EE65
         textCT          =   "frmSales.frx":EE7D
         textRT          =   "frmSales.frx":EE95
         textLM          =   "frmSales.frx":EEAD
         textRM          =   "frmSales.frx":EEC5
         textLB          =   "frmSales.frx":EEDD
         textCB          =   "frmSales.frx":EEF5
         textRB          =   "frmSales.frx":EF0D
         colorBack       =   "frmSales.frx":EF25
         colorIntern     =   "frmSales.frx":EF4F
         colorMO         =   "frmSales.frx":EF79
         colorFocus      =   "frmSales.frx":EFA3
         colorDisabled   =   "frmSales.frx":EFCD
         colorPressed    =   "frmSales.frx":EFF7
         Style           =   2
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdDeptStrip 
         Height          =   1080
         Index           =   4
         Left            =   6525
         TabIndex        =   51
         Top             =   0
         Width           =   1395
         _Version        =   524298
         _ExtentX        =   2461
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "1"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":F021
         textLT          =   "frmSales.frx":F083
         textCT          =   "frmSales.frx":F09B
         textRT          =   "frmSales.frx":F0B3
         textLM          =   "frmSales.frx":F0CB
         textRM          =   "frmSales.frx":F0E3
         textLB          =   "frmSales.frx":F0FB
         textCB          =   "frmSales.frx":F113
         textRB          =   "frmSales.frx":F12B
         colorBack       =   "frmSales.frx":F143
         colorIntern     =   "frmSales.frx":F16D
         colorMO         =   "frmSales.frx":F197
         colorFocus      =   "frmSales.frx":F1C1
         colorDisabled   =   "frmSales.frx":F1EB
         colorPressed    =   "frmSales.frx":F215
         Style           =   2
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdDeptStrip 
         Height          =   1080
         Index           =   6
         Left            =   9315
         TabIndex        =   52
         Top             =   0
         Width           =   1395
         _Version        =   524298
         _ExtentX        =   2461
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "1"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":F23F
         textLT          =   "frmSales.frx":F2A1
         textCT          =   "frmSales.frx":F2B9
         textRT          =   "frmSales.frx":F2D1
         textLM          =   "frmSales.frx":F2E9
         textRM          =   "frmSales.frx":F301
         textLB          =   "frmSales.frx":F319
         textCB          =   "frmSales.frx":F331
         textRB          =   "frmSales.frx":F349
         colorBack       =   "frmSales.frx":F361
         colorIntern     =   "frmSales.frx":F38B
         colorMO         =   "frmSales.frx":F3B5
         colorFocus      =   "frmSales.frx":F3DF
         colorDisabled   =   "frmSales.frx":F409
         colorPressed    =   "frmSales.frx":F433
         Style           =   2
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdDeptStrip 
         Height          =   1080
         Index           =   1
         Left            =   2340
         TabIndex        =   53
         Top             =   0
         Width           =   1395
         _Version        =   524298
         _ExtentX        =   2461
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "1"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":F45D
         textLT          =   "frmSales.frx":F4BF
         textCT          =   "frmSales.frx":F4D7
         textRT          =   "frmSales.frx":F4EF
         textLM          =   "frmSales.frx":F507
         textRM          =   "frmSales.frx":F51F
         textLB          =   "frmSales.frx":F537
         textCB          =   "frmSales.frx":F54F
         textRB          =   "frmSales.frx":F567
         colorBack       =   "frmSales.frx":F57F
         colorIntern     =   "frmSales.frx":F5A9
         colorMO         =   "frmSales.frx":F5D3
         colorFocus      =   "frmSales.frx":F5FD
         colorDisabled   =   "frmSales.frx":F627
         colorPressed    =   "frmSales.frx":F651
         Style           =   2
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdDeptStrip 
         Height          =   1080
         Index           =   2
         Left            =   3735
         TabIndex        =   54
         Top             =   0
         Width           =   1395
         _Version        =   524298
         _ExtentX        =   2461
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "1"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":F67B
         textLT          =   "frmSales.frx":F6DD
         textCT          =   "frmSales.frx":F6F5
         textRT          =   "frmSales.frx":F70D
         textLM          =   "frmSales.frx":F725
         textRM          =   "frmSales.frx":F73D
         textLB          =   "frmSales.frx":F755
         textCB          =   "frmSales.frx":F76D
         textRB          =   "frmSales.frx":F785
         colorBack       =   "frmSales.frx":F79D
         colorIntern     =   "frmSales.frx":F7C7
         colorMO         =   "frmSales.frx":F7F1
         colorFocus      =   "frmSales.frx":F81B
         colorDisabled   =   "frmSales.frx":F845
         colorPressed    =   "frmSales.frx":F86F
         Style           =   2
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdArr 
         Height          =   1080
         Index           =   2
         Left            =   11640
         TabIndex        =   55
         Top             =   0
         Width           =   1785
         _Version        =   524298
         _ExtentX        =   3149
         _ExtentY        =   1905
         _StockProps     =   66
         Caption         =   "Select"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         CornerFactor    =   12
         Surface         =   1
         BackColorContainer=   2134465
         SpecialEffect   =   1
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmSales.frx":F899
         textLT          =   "frmSales.frx":F905
         textCT          =   "frmSales.frx":F91D
         textRT          =   "frmSales.frx":F935
         textLM          =   "frmSales.frx":F94D
         textRM          =   "frmSales.frx":F965
         textLB          =   "frmSales.frx":F97D
         textCB          =   "frmSales.frx":F995
         textRB          =   "frmSales.frx":F9AD
         colorBack       =   "frmSales.frx":F9C5
         colorIntern     =   "frmSales.frx":F9EF
         colorMO         =   "frmSales.frx":FA19
         colorFocus      =   "frmSales.frx":FA43
         colorDisabled   =   "frmSales.frx":FA6D
         colorPressed    =   "frmSales.frx":FA97
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin VSFlex8Ctl.VSFlexGrid grdFind 
         Bindings        =   "frmSales.frx":FAC1
         Height          =   10050
         Left            =   0
         TabIndex        =   69
         Top             =   1230
         Width           =   13695
         _cx             =   24156
         _cy             =   17727
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
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
         ForeColor       =   -2147483640
         BackColorFixed  =   7848417
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8963553
         ForeColorSel    =   0
         BackColorBkg    =   16777215
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
         Rows            =   50
         Cols            =   7
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
         WallPaper       =   "frmSales.frx":FAD7
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSForms.Image Image1 
         Height          =   165
         Left            =   -90
         Top             =   1080
         Width           =   14865
         BackColor       =   2267338
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "26220;291"
      End
      Begin MSForms.Image Image3 
         Height          =   1215
         Left            =   13410
         Top             =   -90
         Width           =   165
         BackColor       =   2267338
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "291;2143"
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00B4DAED&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0086C4E1&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0071B9DB&
         FillStyle       =   2  'Horizontal Line
         Height          =   8175
         Left            =   13710
         Top             =   1980
         Width           =   1065
      End
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1965
      Index           =   5
      Left            =   375
      TabIndex        =   67
      ToolTipText     =   " <Ctrl-R> "
      Top             =   4305
      Width           =   1965
      _Version        =   524298
      _ExtentX        =   3466
      _ExtentY        =   3466
      _StockProps     =   66
      Caption         =   "R/A"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Shape           =   6
      CornerFactor    =   15
      BackColorContainer=   14737632
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales.frx":13808
      textLT          =   "frmSales.frx":1386E
      textCT          =   "frmSales.frx":13886
      textRT          =   "frmSales.frx":1389E
      textLM          =   "frmSales.frx":138B6
      textRM          =   "frmSales.frx":138CE
      textLB          =   "frmSales.frx":138E6
      textCB          =   "frmSales.frx":138FE
      textRB          =   "frmSales.frx":13916
      colorBack       =   "frmSales.frx":1392E
      colorIntern     =   "frmSales.frx":13958
      colorMO         =   "frmSales.frx":13982
      colorFocus      =   "frmSales.frx":139AC
      colorDisabled   =   "frmSales.frx":139D6
      colorPressed    =   "frmSales.frx":13A00
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1125
      Index           =   2
      Left            =   2490
      TabIndex        =   61
      ToolTipText     =   " <Ctrl-Y> "
      Top             =   4020
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      Caption         =   "Pay Out"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales.frx":13A2A
      textLT          =   "frmSales.frx":13A98
      textCT          =   "frmSales.frx":13AB0
      textRT          =   "frmSales.frx":13AC8
      textLM          =   "frmSales.frx":13AE0
      textRM          =   "frmSales.frx":13AF8
      textLB          =   "frmSales.frx":13B10
      textCB          =   "frmSales.frx":13B28
      textRB          =   "frmSales.frx":13B40
      colorBack       =   "frmSales.frx":13B58
      colorIntern     =   "frmSales.frx":13B82
      colorMO         =   "frmSales.frx":13BAC
      colorFocus      =   "frmSales.frx":13BD6
      colorDisabled   =   "frmSales.frx":13C00
      colorPressed    =   "frmSales.frx":13C2A
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdSlip 
      Height          =   1755
      Left            =   390
      TabIndex        =   59
      ToolTipText     =   " <Ctrl-P> "
      Top             =   2280
      Width           =   2120
      _Version        =   524298
      _ExtentX        =   3739
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Slip on"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales.frx":13C54
      textLT          =   "frmSales.frx":13CC2
      textCT          =   "frmSales.frx":13CDA
      textRT          =   "frmSales.frx":13CF2
      textLM          =   "frmSales.frx":13D0A
      textRM          =   "frmSales.frx":13D22
      textLB          =   "frmSales.frx":13D3A
      textCB          =   "frmSales.frx":13D52
      textRB          =   "frmSales.frx":13D6A
      colorBack       =   "frmSales.frx":13D82
      colorIntern     =   "frmSales.frx":13DAC
      colorMO         =   "frmSales.frx":13DD6
      colorFocus      =   "frmSales.frx":13E00
      colorDisabled   =   "frmSales.frx":13E2A
      colorPressed    =   "frmSales.frx":13E54
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   3
      Left            =   2490
      TabIndex        =   57
      ToolTipText     =   " <Ctrl-B> "
      Top             =   2280
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Reprint"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
      textCaption     =   "frmSales.frx":13E7E
      textLT          =   "frmSales.frx":13EEC
      textCT          =   "frmSales.frx":13F04
      textRT          =   "frmSales.frx":13F1C
      textLM          =   "frmSales.frx":13F34
      textRM          =   "frmSales.frx":13F4C
      textLB          =   "frmSales.frx":13F64
      textCB          =   "frmSales.frx":13F7C
      textRB          =   "frmSales.frx":13F94
      colorBack       =   "frmSales.frx":13FAC
      colorIntern     =   "frmSales.frx":13FD6
      colorMO         =   "frmSales.frx":14000
      colorFocus      =   "frmSales.frx":1402A
      colorDisabled   =   "frmSales.frx":14054
      colorPressed    =   "frmSales.frx":1407E
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1095
      Index           =   8
      Left            =   2505
      TabIndex        =   66
      ToolTipText     =   " <F3> "
      Top             =   9475
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   2265040
      Caption         =   "Voucher"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1155
      Index           =   7
      Left            =   2505
      TabIndex        =   65
      ToolTipText     =   " <F4> "
      Top             =   8340
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2037
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   2265040
      Caption         =   "Charge"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1125
      Index           =   6
      Left            =   2505
      TabIndex        =   64
      ToolTipText     =   " <F2> "
      Top             =   7245
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1984
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   2265040
      Caption         =   "Card"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1125
      Index           =   4
      Left            =   2490
      TabIndex        =   62
      Top             =   5070
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1984
      _StockProps     =   66
      Caption         =   "Discount"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
      BackColorContainer=   14737632
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales.frx":140A8
      textLT          =   "frmSales.frx":14118
      textCT          =   "frmSales.frx":14130
      textRT          =   "frmSales.frx":14148
      textLM          =   "frmSales.frx":14160
      textRM          =   "frmSales.frx":14178
      textLB          =   "frmSales.frx":14190
      textCB          =   "frmSales.frx":141A8
      textRB          =   "frmSales.frx":141C0
      colorBack       =   "frmSales.frx":141D8
      colorIntern     =   "frmSales.frx":14202
      colorMO         =   "frmSales.frx":1422C
      colorFocus      =   "frmSales.frx":14256
      colorDisabled   =   "frmSales.frx":14280
      colorPressed    =   "frmSales.frx":142AA
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.Timer scrolTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   200
      Top             =   960
   End
   Begin VSFlex8Ctl.VSFlexGrid grdDept 
      Height          =   8070
      Left            =   -30
      TabIndex        =   47
      Top             =   780
      Visible         =   0   'False
      Width           =   135
      _cx             =   238
      _cy             =   14235
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
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.Timer voidTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4260
      Top             =   570
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   825
      Left            =   4620
      TabIndex        =   37
      Top             =   2190
      Visible         =   0   'False
      Width           =   9525
      _Version        =   524298
      _ExtentX        =   16801
      _ExtentY        =   1455
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
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      UserData        =   0.1
      textCaption     =   "frmSales.frx":142D4
      textLT          =   "frmSales.frx":1435A
      textCT          =   "frmSales.frx":14372
      textRT          =   "frmSales.frx":1438A
      textLM          =   "frmSales.frx":143A2
      textRM          =   "frmSales.frx":143BA
      textLB          =   "frmSales.frx":143D2
      textCB          =   "frmSales.frx":143EA
      textRB          =   "frmSales.frx":14402
      colorBack       =   "frmSales.frx":1441A
      colorIntern     =   "frmSales.frx":14444
      colorMO         =   "frmSales.frx":1446E
      colorFocus      =   "frmSales.frx":14498
      colorDisabled   =   "frmSales.frx":144C2
      colorPressed    =   "frmSales.frx":144EC
   End
   Begin VB.PictureBox picDigit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   785
      Left            =   13320
      Picture         =   "frmSales.frx":14516
      ScaleHeight     =   780
      ScaleWidth      =   525
      TabIndex        =   38
      Top             =   2220
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
         Left            =   60
         TabIndex        =   39
         Top             =   150
         Width           =   405
      End
   End
   Begin VB.Timer errTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3780
      Top             =   570
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3240
      Top             =   540
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4470
      Top             =   10770
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7185
      Left            =   4710
      ScaleHeight     =   7155
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   3270
      Width           =   6045
      Begin VSFlex8Ctl.VSFlexGrid grdMain 
         Height          =   7140
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5355
         _cx             =   9446
         _cy             =   12594
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
         GridColorFixed  =   -2147483646
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
         WallPaper       =   "frmSales.frx":1484F
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   600
         Index           =   0
         Left            =   5310
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   -30
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1058
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
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   600
         Index           =   1
         Left            =   5310
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   6570
         Width           =   735
         _ExtentX        =   1296
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
      Begin VB.Shape picGrid 
         BackColor       =   &H00B4DAED&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0086C4E1&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0071B9DB&
         FillStyle       =   2  'Horizontal Line
         Height          =   7125
         Left            =   5340
         Top             =   0
         Width           =   705
      End
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1035
      Index           =   4
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "F9"
      Top             =   10200
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1826
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R10-00"
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
      MaskColor       =   4
      MaskColor       =   4
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1005
      Index           =   3
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "F10"
      Top             =   9210
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R20-00"
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
      MaskColor       =   3
      MaskColor       =   3
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1005
      Index           =   2
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "F11"
      Top             =   8220
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R50-00"
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
      MaskColor       =   2
      MaskColor       =   2
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1005
      Index           =   1
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "F12"
      Top             =   7230
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R100-00"
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
      MaskColor       =   1
      MaskColor       =   1
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1005
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   6270
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "R200-00"
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
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   960
      Index           =   0
      Left            =   10830
      TabIndex        =   11
      Top             =   6480
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1693
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
      Height          =   960
      Index           =   1
      Left            =   12240
      TabIndex        =   12
      Top             =   6480
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1693
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
      Height          =   960
      Index           =   2
      Left            =   13680
      TabIndex        =   13
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1693
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
      Height          =   960
      Index           =   3
      Left            =   10830
      TabIndex        =   14
      Top             =   7470
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1693
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
      Height          =   960
      Index           =   4
      Left            =   12240
      TabIndex        =   15
      Top             =   7470
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1693
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
      Height          =   960
      Index           =   5
      Left            =   13680
      TabIndex        =   16
      Top             =   7470
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1693
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
      Height          =   960
      Index           =   6
      Left            =   10830
      TabIndex        =   17
      Top             =   8460
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1693
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
      Height          =   960
      Index           =   7
      Left            =   12240
      TabIndex        =   18
      Top             =   8460
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1693
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
      Height          =   960
      Index           =   8
      Left            =   13680
      TabIndex        =   19
      Top             =   8460
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1693
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
      Height          =   990
      Index           =   9
      Left            =   10830
      TabIndex        =   20
      Tag             =   " <Back Space> "
      Top             =   9450
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1746
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
      Height          =   990
      Index           =   10
      Left            =   12240
      TabIndex        =   21
      Top             =   9450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
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
      Height          =   990
      Index           =   11
      Left            =   13680
      TabIndex        =   22
      Top             =   9450
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1746
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
      Height          =   990
      Index           =   12
      Left            =   10830
      TabIndex        =   23
      ToolTipText     =   " <*> "
      Top             =   5370
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "x"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   990
      Index           =   13
      Left            =   12240
      TabIndex        =   24
      ToolTipText     =   " <.> "
      Top             =   5370
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   990
      Index           =   14
      Left            =   10830
      TabIndex        =   25
      Top             =   3270
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "No Sale"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   990
      Index           =   15
      Left            =   12240
      TabIndex        =   26
      ToolTipText     =   " <F7> "
      Top             =   3270
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Void All"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   990
      Index           =   16
      Left            =   13680
      TabIndex        =   27
      ToolTipText     =   " <F8> "
      Top             =   3270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Return Item"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   990
      Index           =   17
      Left            =   10830
      TabIndex        =   28
      ToolTipText     =   " <Ctrl-Q> "
      Top             =   4320
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Quote"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   990
      Index           =   18
      Left            =   12240
      TabIndex        =   29
      ToolTipText     =   " <Ctrl-S> "
      Top             =   4320
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Search"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   990
      Index           =   19
      Left            =   13680
      TabIndex        =   30
      ToolTipText     =   " <Enter> "
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Plu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   990
      Index           =   20
      Left            =   13680
      TabIndex        =   31
      ToolTipText     =   " <+> "
      Top             =   5370
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1746
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "Price O/V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdNext 
      Height          =   910
      Left            =   14280
      TabIndex        =   32
      Top             =   2130
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1614
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
   Begin VSFlex8Ctl.VSFlexGrid grdMenu 
      Height          =   5385
      Left            =   120
      TabIndex        =   56
      Top             =   240
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
   Begin BTNENHLib4.BtnEnh cmdAppro 
      Height          =   1755
      Left            =   390
      TabIndex        =   58
      ToolTipText     =   " <Ctrl-A> "
      Top             =   2280
      Visible         =   0   'False
      Width           =   1245
      _Version        =   524298
      _ExtentX        =   2196
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Appro's"
      Enabled         =   0   'False
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
      Shape           =   1
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales.frx":166F2
      textLT          =   "frmSales.frx":16760
      textCT          =   "frmSales.frx":16778
      textRT          =   "frmSales.frx":16790
      textLM          =   "frmSales.frx":167A8
      textRM          =   "frmSales.frx":167C0
      textLB          =   "frmSales.frx":167D8
      textCB          =   "frmSales.frx":167F0
      textRB          =   "frmSales.frx":16808
      colorBack       =   "frmSales.frx":16820
      colorIntern     =   "frmSales.frx":1684A
      colorMO         =   "frmSales.frx":16874
      colorFocus      =   "frmSales.frx":1689E
      colorDisabled   =   "frmSales.frx":168C8
      colorPressed    =   "frmSales.frx":168F2
      Orientation     =   5
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
      Left            =   570
      ScaleHeight     =   615
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   2430
      Width           =   825
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   2145
      Index           =   1
      Left            =   360
      TabIndex        =   60
      ToolTipText     =   " <Ctrl-U> "
      Top             =   4020
      Width           =   2145
      _Version        =   524298
      _ExtentX        =   3784
      _ExtentY        =   3784
      _StockProps     =   66
      Caption         =   "Wastage"
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
      Shape           =   6
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSales.frx":1691C
      textLT          =   "frmSales.frx":1698A
      textCT          =   "frmSales.frx":169A2
      textRT          =   "frmSales.frx":169BA
      textLM          =   "frmSales.frx":169D2
      textRM          =   "frmSales.frx":169EA
      textLB          =   "frmSales.frx":16A02
      textCB          =   "frmSales.frx":16A1A
      textRB          =   "frmSales.frx":16A32
      colorBack       =   "frmSales.frx":16A4A
      colorIntern     =   "frmSales.frx":16A74
      colorMO         =   "frmSales.frx":16A9E
      colorFocus      =   "frmSales.frx":16AC8
      colorDisabled   =   "frmSales.frx":16AF2
      colorPressed    =   "frmSales.frx":16B1C
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1095
      Index           =   5
      Left            =   2505
      TabIndex        =   63
      ToolTipText     =   " <F1> "
      Top             =   6195
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   2265040
      Caption         =   "Cash"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdLogOff 
      Height          =   1560
      Left            =   360
      TabIndex        =   68
      Top             =   360
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   2752
      Appearance      =   3
      BackColor       =   2720171
      Caption         =   "Log Off"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin MSAdodcLib.Adodc adoData 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      ConnectMode     =   1
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoData"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSForms.Label lblDebtor 
      Height          =   285
      Left            =   6510
      TabIndex        =   73
      Top             =   10875
      Width           =   4125
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7276;503"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label lblHappy1 
      Height          =   345
      Left            =   1320
      TabIndex        =   72
      Top             =   1920
      Visible         =   0   'False
      Width           =   2265
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Happy Hour Active"
      Size            =   "3995;609"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label lblPayReason 
      Height          =   495
      Left            =   5610
      TabIndex        =   71
      Top             =   810
      Visible         =   0   'False
      Width           =   1365
   End
   Begin MSForms.Label lblHappy 
      Height          =   345
      Left            =   1320
      TabIndex        =   70
      Top             =   1940
      Visible         =   0   'False
      Width           =   2265
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Happy Hour Active"
      Size            =   "3995;609"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblTender 
      Height          =   1605
      Left            =   6570
      TabIndex        =   35
      Top             =   420
      Width           =   8235
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "14526;2831"
      FontName        =   "Calibri"
      FontHeight      =   1320
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblKeyRegister 
      Height          =   645
      Left            =   4770
      TabIndex        =   34
      Top             =   2340
      Width           =   9100
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "16051;1138"
      FontName        =   "Calibri"
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblUser 
      Height          =   285
      Left            =   10710
      TabIndex        =   33
      Top             =   10875
      Width           =   4095
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7223;503"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCash 
      Height          =   1545
      Left            =   2520
      TabIndex        =   10
      Top             =   330
      Width           =   7845
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "13838;2725"
      FontName        =   "Calibri"
      FontHeight      =   1320
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape shpLive 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   0
      Left            =   570
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   165
   End
   Begin VB.Shape shpLive 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   1
      Left            =   825
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   165
   End
   Begin VB.Shape shpLive 
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   2
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   165
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   2730
      TabIndex        =   9
      Top             =   10875
      Width           =   3705
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "6535;503"
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Image newBack 
      Height          =   1260
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "2037;2222"
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAppro_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Finalizing = True Then Exit Sub
    Keystring = "Appros"
    If TillData.ShortTender = True Then
        DisplayErr "Invalid Key Pressed"
        Exit Sub
    End If
    Select Case cmdAppro.Caption
        Case "Fuel Pumps"
            Load frmPumps
            frmPumps.top = frmSales.Picture1.top - 150
            frmPumps.Left = frmSales.Picture1.Left - 100
            frmPumps.Show vbModal
        Case "PDT Upload"
            PDTSale
        Case Else
            frmAppro.Show vbModal
    End Select
End Sub
Private Sub PDTSale()
    frmSales.Tag = "1"
    On Error Resume Next
    Screen.MousePointer = 11
    stime = 0
    Kill App.Path & "\Cipher\Test.txt"
    DoEvents
    Shell App.Path & "\Cipher\Data_Read.exe", vbNormalFocus
    DoEvents
    stime = Timer
    While Timer - stime < 12: Wend
    On Error GoTo trap
    filenum = FreeFile
    Open App.Path & "\Cipher\Test.txt" For Input As filenum
    i = 0
    picHoldFocus.SetFocus
    Barcode = ""
    While Not EOF(filenum)
        DoEvents
        Line Input #filenum, newline
        Barcode = Trim(Mid(newline, 1, InStr(newline, ",") - 1))
        Qty = Val(Trim(Mid(newline, InStr(newline, ",") + 1)))
        PdtString = ""
        PdtString = Trim(Str(Qty) & "*" & Trim(Barcode))
        For i = 1 To Len(PdtString)
            SendKeys Mid(PdtString, i, 1)
            picHoldFocus.SetFocus
        Next i
        SendKeys "{ENTER}"
        picHoldFocus.SetFocus
    Wend
    Close filenum
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
trap:
    Close filenum
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
End Sub
Private Sub cmdArr_Click(Index As Integer)
    Select Case Index
        Case 1
            If grdDept.Row = grdDept.Rows - 1 Then Exit Sub
            b = -1
            For i = grdDept.Row + 1 To grdDept.Row + 7
                b = b + 1
                If i < grdDept.Rows Then
                    cmdDeptStrip(b).Caption = Replace(grdDept.TextMatrix(i, 0), "&", "&&")
                    cmdDeptStrip(b).Tag = grdDept.TextMatrix(i, 1)
                    If grdDept.TextMatrix(i, 2) = "1" Then
                        cmdDeptStrip(b).Value = 1
                    Else
                        cmdDeptStrip(b).Value = 0
                    End If
                    grdDept.Row = grdDept.Row + 1
                    If cmdDeptStrip(b).Visible = False Then cmdDeptStrip(b).Visible = True
                Else
                    cmdDeptStrip(b).Visible = False
                End If
            Next i
        Case 0
            If grdDept.Row = 6 Then Exit Sub
            b = -1
            extr = 0
            If cmdDeptStrip(6).Visible = False Then
                For i = 0 To 6
                    If cmdDeptStrip(i).Visible = False Then extr = extr + 1
                Next i
            End If
            For i = grdDept.Row - 13 + extr To grdDept.Row - 7 + extr
                b = b + 1
                If i < grdDept.Rows Then
                    cmdDeptStrip(b).Caption = Replace(grdDept.TextMatrix(i, 0), "&", "&&")
                    cmdDeptStrip(b).Tag = grdDept.TextMatrix(i, 1)
                    If grdDept.TextMatrix(i, 2) = "1" Then
                        cmdDeptStrip(b).Value = 1
                    Else
                        cmdDeptStrip(b).Value = 0
                    End If
                    grdDept.Row = grdDept.Row - 1
                    If cmdDeptStrip(b).Visible = False Then cmdDeptStrip(b).Visible = True
                Else
                    cmdDeptStrip(b).Visible = False
                End If
            Next i
            If extr <> 0 Then grdDept.Row = grdDept.Row + extr
        Case 2
            cmdArr(2).Tag = grdFind.TextMatrix(grdFind.Row, 0)
            If Panel_no = 0 Then frmSales.KeyPreview = True
            picSearch.Visible = False
            If InStr(KeyRegister, Chr$(215) & " *") <> 0 Then
                lblKeyRegister.Caption = Mid(lblKeyRegister.Caption, 1, InStr(lblKeyRegister.Caption, Chr$(215)) + 1)
            End If
            If InStr(KeyRegister, " (Return Item) ") <> 0 Then
                If InStr(KeyRegister, Chr$(215)) = 0 Then
                    lblKeyRegister = " (Return Item) "
                End If
            End If
            KeyRegister = lblKeyRegister.Caption & cmdArr(2).Tag
            Key_Function "Plu"
            picHoldFocus.SetFocus
            Exit Sub
    End Select
    grdFind.SetFocus
End Sub

Private Sub cmdArrow_Click(Index As Integer)
    If grdMain.Rows = 1 Then Exit Sub
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    
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
            grdMain.ShowCell grdMain.Row - 1, 0
    End Select
End Sub

Private Sub cmdArrow1_Click(Index As Integer)
    Select Case Index
        Case 0
            If grdFind.Row > 1 Then
                grdFind.Row = grdFind.Row - 1
            End If
            grdFind.ShowCell grdFind.Row, 0
        Case 1
            If grdFind.Row < grdFind.Rows - 1 Then
                grdFind.Row = grdFind.Row + 1
            End If
            grdFind.ShowCell grdFind.Row, 0
    End Select
    grdFind.SetFocus
End Sub

Private Sub cmdArrow1_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    scrolTimer.Tag = Index
    scrolTimer.Enabled = True
End Sub

Private Sub cmdArrow1_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    scrolTimer.Enabled = False
    scrolTimer.Interval = 1000
End Sub

Private Sub cmdClose_Click()
    If Panel_no = 0 Then frmSales.KeyPreview = True
    KeyRegister = ""
    picSearch.Visible = False
End Sub

Private Sub cmdDeptStrip_Click(Index As Integer)
    For i = 0 To grdDept.Rows - 1
        If cmdDeptStrip(Index).Tag = grdDept.TextMatrix(i, 1) Then
            grdDept.TextMatrix(i, 2) = "1"
        Else
            grdDept.TextMatrix(i, 2) = ""
        End If
    Next i
    adoData.ConnectionString = cnnMain.ConnectionString
    adoData.CursorLocation = adUseServer
    adoData.CursorType = adOpenStatic
    adoData.LockType = adLockReadOnly
    adoData.RecordSource = "Select Product_Code,Description,Department,SOH,Landed_Cost,Tax_Rate,Selling_Price from Product_List where Sales_Item=1 and Department_No = '" & cmdDeptStrip(Index).Tag & "' order by Description"
    adoData.Refresh
    grdFind.Col = 1
    If grdFind.Rows > 1 Then grdFind.Row = 1
    picKey.top = 2010
    grdFind.SetFocus
End Sub
Private Sub cmdErr_Click()
    If grdMain.Enabled = False Then
        send_data_steam_keylog (Me.Name & " - " & cmdErr(Index).Caption)
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    cmdInput(15).Caption = "Void All"
    cmdErr.Visible = False
    cmdErr.BackColor = &HF2&
    errTimer.Enabled = False
    KeyRegister = ""
    lblKeyRegister = ""
    picDigit.Visible = False
End Sub

Private Sub cmdFancy_Click(Index As Integer)
    On Error Resume Next
    If Finalizing = True Then Exit Sub
    send_data_steam_keylog (Me.Name & " - " & cmdInput(Index).Caption)
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    If cmdErr.Visible = True Then Exit Sub
    If cmdFancy(Index).Caption <> "Reprint" Then Reprintplease = False: Thediscounttotal = 0
    If cmdFancy(Index).Caption = "Reprint" Then Reprintplease = True:
    If Reprintplease = False Then Thediscounttotal = 0
    Key_Function cmdFancy(Index).Caption
    
    
    If picHoldFocus.Tag = "" Then
        picHoldFocus.SetFocus
    Else
        picHoldFocus.Tag = ""
    End If
    cmdInput(15).Caption = "Void All"
    On Error GoTo 0
End Sub

Private Sub cmdInput_Click(Index As Integer)
    Thediscounttotal = 0
    If Finalizing = True Then Exit Sub
    send_data_steam_keylog (Me.Name & " - " & cmdInput(Index).Caption)
        Select Case Index
            Case 12 To 16, 18, 19, 20, 9
                cmdKey(7).Enabled = True
    
            Case 0 To 8, 10, 11
                cmdKey(7).Enabled = False
    
            ' Quote
            Case 17
                cmdKey(7).Enabled = True
    
                If UserRecord.Quotes = False Then
                    frmSales.Tag = "1"
                    TillData.UserOveride = 0
                    Load frmValidate
                    frmValidate.Tag = "Quotes"
                    frmValidate.Show vbModal
                    frmSales.Tag = ""
                    Select Case frmValidate.Tag
                        Case "0"
                            frmValidate.Tag = ""
                            cmdErr.Caption = "Higher Access Rights required to Quote"
                            cmdErr.Visible = True
                            errTimer.Enabled = True
                            picHoldFocus.SetFocus
                            Exit Sub
                        Case ""
                            frmValidate.Tag = ""
                            KeyRegister = ""
                            Exit Sub
                        Case "1"
                            frmValidate.Tag = ""
                    End Select
                End If
    
        End Select
    
        If grdMain.Enabled = False And cmdInput(Index).Caption <> "Void" Then
            voidTimer.Enabled = False
            grdMain.Enabled = True
            grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
            grdMain.HighLight = flexHighlightAlways
        End If
    
        'xxxxxxxxxxxxxxx1
        If cmdInput(Index).Caption = "CL" Then
            frmSales.cmdErr.Visible = False
            cmdErr.BackColor = &HF2&
            errTimer.Enabled = False
            cmdInput(15).Caption = "Void All"
            picHoldFocus.SetFocus
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
                Exit Sub
            End If
        End If
        If cmdInput(Index).Caption = "Void All" Then grdMain.Enabled = False
        If grdMain.Enabled = False And (cmdInput(Index).Caption = "Void" Or cmdInput(Index).Caption = "Void All") Then
            voidTimer.Enabled = False
            grdMain.Enabled = True
            grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
            grdMain.HighLight = flexHighlightAlways
            If grdMain.TextMatrix(grdMain.Row, 9) <> "0" Then
                If UserRecord.Voids = False Then
                    frmSales.Tag = "1"
                    TillData.UserOveride = 0
                    Load frmValidate
                    frmValidate.Tag = "Void"
                    frmValidate.Show vbModal
                    frmSales.Tag = ""
                    Select Case frmValidate.Tag
                        Case "0"
                            frmValidate.Tag = ""
                            cmdErr.Caption = "Higher Access Rights required to Void"
                            cmdErr.Visible = True
                            errTimer.Enabled = True
                            picHoldFocus.SetFocus
                            Exit Sub
                        Case ""
                            frmValidate.Tag = ""
                            KeyRegister = ""
                            Exit Sub
                        Case "1"
                            frmValidate.Tag = ""
                    End Select
                End If
                If cmdInput(Index).Caption = "Void All" Then 'Kotie 18-03-2013 07:50
                    For v = 1 To grdMain.Rows - 1
                        If grdMain.TextMatrix(v, 2) <> "" And grdMain.TextMatrix(v, 0) <> "" Then
                            If grdMain.Cell(flexcpForeColor, v, 0, v, 2) <> vbRed Then
                                If InStr(grdMain.TextMatrix(v, 8), "Void") = 0 Then
                                    LineToVoid = v
                                    If Val(grdMain.TextMatrix(v, 0)) = 0 Then Qty = 1 Else Qty = Val(grdMain.TextMatrix(v, 0))
                                    KeyRegister = " (Void) " & grdMain.TextMatrix(v, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(v, 2)) / Qty) * 100 & " (Price O/V) " & grdMain.TextMatrix(v, 9)
                                    Qty = Null
                                    Key_Function "Plu"
                                    LineToVoid = 0
                                End If
                            End If
                        End If
                    Next v
                    TillData.Discount = 0
                    TillData.DiscountVal = 0
                    TillData.TotDiscount = 0
                    TillData.TotDiscountCount = 0
                    TillData.TotDiscountVal = 0
                    TillData.TotDiscountValCount = 0
                Else
                    LineToVoid = grdMain.Row
                    KeyRegister = " (Void) " & grdMain.TextMatrix(grdMain.Row, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(grdMain.Row, 2)) / Val(grdMain.TextMatrix(grdMain.Row, 0))) * 100 & " (Price O/V) " & grdMain.TextMatrix(grdMain.Row, 9)
                    Key_Function "Plu"
                    LineToVoid = 0
                End If
                cmdInput(15).Caption = "Void All"
            'KeyRegister = " (Void) " & grdMain.TextMatrix(grdMain.Row, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(grdMain.Row, 2)) / Val(grdMain.TextMatrix(grdMain.Row, 0))) * 100 & " (Price O/V) " & grdMain.TextMatrix(grdMain.Row, 9)
            'Key_Function "Plu"
            Else
                If cmdInput(Index).Caption = "Void All" Then  'Kotie 18-03-2013 07:50
                    For v = grdMain.Rows - 1 To 1 Step -1
                        If grdMain.TextMatrix(v, 2) <> "" Then
                            LineToVoid = v
                            If Val(grdMain.TextMatrix(v, 0)) = 0 Then Qty = 1 Else Val (grdMain.TextMatrix(v, 0))
                            KeyRegister = " (Void) " & grdMain.TextMatrix(v, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(v, 2)) / Qty) * 100 & " (Price O/V) " & grdMain.TextMatrix(v, 9)
                            Qty = Null
                            Key_Function "Plu"
                        End If
                    Next v
                Else
                    KeyRegister = " (Void) " & grdMain.TextMatrix(grdMain.Row, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(grdMain.Row, 2)) / Val(grdMain.TextMatrix(grdMain.Row, 0))) * 100 & " (Price O/V) " & grdMain.TextMatrix(grdMain.Row, 9)
                    Key_Function "Plu"
                End If
            'KeyRegister = " (Void) " & grdMain.TextMatrix(grdMain.Row, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(grdMain.Row, 2)) / Val(grdMain.TextMatrix(grdMain.Row, 0))) * 100 & " (Price O/V) " & grdMain.TextMatrix(grdMain.Row, 9)
            'Key_Function "Plu"
            End If
            picHoldFocus.SetFocus
            Exit Sub
        End If
        If cmdErr.Visible = True Then Exit Sub
        Key_Function cmdInput(Index).Caption
        If picSearch.Visible = True Then
            grdFind.SetFocus
        Else
            If frmSales.Visible Then
            picHoldFocus.SetFocus
        End If
    End If
    
End Sub
Private Sub cmdKey_Click(Index As Integer)
    
    Screen.MousePointer = 1
    On Error Resume Next
    If Finalizing = True Then Exit Sub
    send_data_steam_keylog (Me.Name & " - " & cmdKey(Index).Caption)
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    If cmdErr.Visible = True Then Exit Sub
    
    Select Case Index
    Case 0 To 8
    cmdKey(0).Enabled = False
    cmdKey(1).Enabled = False
    cmdKey(2).Enabled = False
    cmdKey(3).Enabled = False
    cmdKey(4).Enabled = False
    cmdKey(5).Enabled = False
    cmdKey(6).Enabled = False
    cmdKey(7).Enabled = False
    cmdKey(8).Enabled = False
    
    End Select
    
    Key_Function cmdKey(Index).Caption
    
    If picHoldFocus.Tag = "" Then
        picHoldFocus.SetFocus
        
    Else
        picHoldFocus.Tag = ""
    End If
    cmdKey(0).Enabled = True
    cmdKey(1).Enabled = True
    cmdKey(2).Enabled = True
    cmdKey(3).Enabled = True
    cmdKey(4).Enabled = True
    cmdKey(5).Enabled = True
    cmdKey(6).Enabled = True
    cmdKey(7).Enabled = True
    cmdKey(8).Enabled = True
    On Error GoTo 0
End Sub
Private Sub cmdKeyboard_Click()
    If grdFind.Rows = 1 Then Exit Sub
    On Error Resume Next
    cmdKeyboard.SpecialEffect = fmSpecialEffectSunken
    Screen.MousePointer = 11
    DoEvents
    frmSales.Tag = "1"
    Load frmKeyBoard
    frmKeyBoard.Tag = "frmSales"
    Screen.MousePointer = 0
    frmKeyBoard.Show vbModal
    cmdKeyboard.SpecialEffect = fmSpecialEffectFlat
    cmdKeyboard.BorderStyle = fmBorderStyleSingle
    grdFind.SetFocus
    DoEvents
    On Error GoTo 0
End Sub
Private Sub cmdLogoff_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
        frmSales.cmdErr.Caption = "Finalize the Sale before Logging Off"
        frmSales.cmdErr.Visible = True
        frmSales.errTimer.Enabled = True
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
    frmSales.KeyPreview = False
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
    grdMain.Rows = 1
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
                    Case "Reservations"
                        frmMain.cmdBar(0).Enabled = False
                        frmRes.Show
                    Case "Rooms"
                        frmMain.cmdBar(1).Enabled = False
                        frmRooms.Show
                    Case "Guests"
                        frmMain.cmdBar(2).Enabled = False
                        frmGuests.Show
                    Case "Checkin"
                        frmMain.cmdBar(3).Enabled = False
                        frmCheck.Show
                    Case "Stock"
                        frmMain.cmdBar(7).Enabled = False
                        frmProducts.Show
                    Case "Users"
                        frmMain.cmdBar(4).Enabled = False
                        frmUsers.Show
                    Case "Reports"
                        frmMain.cmdBar(5).Enabled = False
                        frmReports.Show
                    Case Else
                        frmDetails.Show
                End Select
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
                    If Form.Name = "frmSales1" Then Unload frmSales1
                    If Form.Name = "frmRestRes" Then Unload frmRestRes
                    If Form.Name = "frmTillReport" Then Unload frmTillReport
                    If Form.Name = "frmBar" Then Unload frmBar
                    Next Form
            Unload Me
        Case Else
            Screen.MousePointer = 1
            frmMain.cmdBar(6).Enabled = True
            frmMain.Show
            Select Case frmSales.Tag
                Case "Reservations"
                    frmMain.cmdBar(0).Enabled = False
                    frmRes.Show
                Case "Rooms"
                    frmMain.cmdBar(1).Enabled = False
                    frmRooms.Show
                Case "Guests"
                    frmMain.cmdBar(2).Enabled = False
                    frmGuests.Show
                Case "Checkin"
                    frmMain.cmdBar(3).Enabled = False
                    frmCheck.Show
                Case "Stock"
                    frmMain.cmdBar(7).Enabled = False
                    frmProducts.Load_Products
                    frmProducts.Show
                Case "Users"
                    frmMain.cmdBar(4).Enabled = False
                    frmUsers.Show
                Case "Reports"
                    frmMain.cmdBar(5).Enabled = False
                    'frmReports.Show
                    frmReports.WindowState = 2
                Case Else
                    frmDetails.Show
            End Select
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
Private Sub cmdNext_Click()
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    If cmdErr.Visible = True Then Exit Sub
    If UserRecord.uType = 4 Or UserRecord.uType = 8 Then
        With frmBar
            Screen.MousePointer = 11
            picHoldFocus.SetFocus
            frmBar.Show
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
            If TillData.DocNo = 0 Then
                .cmdKey(4).Caption = "No Sale"
                .grdMain.HighLight = flexHighlightWithFocus
            Else
                .cmdKey(4).Caption = "Corr"
                .grdMain.HighLight = flexHighlightAlways
               
            End If
            .grdMain.Row = grdMain.Row
            .lblKeyRegister.Caption = lblKeyRegister.Caption
            .lblCash.Caption = lblCash.Caption
            .lblTender = lblTender.Caption
            .grdMain.TopRow = grdMain.TopRow
        End With
    Else
         With frmSales1
            Screen.MousePointer = 11
            picHoldFocus.SetFocus
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
            If TillData.DocNo = 0 Then
                .cmdDept(6).Caption = "No Sale"
                .grdMain.HighLight = flexHighlightWithFocus
            Else
                .cmdDept(6).Caption = "Corr"
                .grdMain.HighLight = flexHighlightAlways
            End If
            .grdMain.Row = grdMain.Row
            .lblKeyRegister.Caption = lblKeyRegister.Caption
            .lblCash.Caption = lblCash.Caption
            .lblTender = lblTender.Caption
            .grdMain.TopRow = grdMain.TopRow
        End With
    End If
End Sub

Private Sub cmdSlip_Click()
    send_data_steam_keylog (Me.Name & " - " & cmdSlip(Index).Caption)
    Select Case cmdSlip.Caption
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
End Sub

Private Sub errTimer_Timer()
    Select Case cmdErr.BackColor
        Case &HF2&      'White
            cmdErr.BackColor = &H44B0DF
        Case &H44B0DF       'Yellow
            cmdErr.BackColor = &HF2&
    End Select
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    chkdateagain = ""
    picHoldFocus.SetFocus
    If grdMain.Rows = 1 Then
        cmdFancy(4).Caption = "Member No"
    Else
        cmdFancy(4).Caption = "Discount"
    End If
    ActiveReadServer "Select * from Debtors where Debtor_No ='" & TillData.Account_No & "'"
    If rs.RecordCount > 0 Then
        frmSales.lblKeyRegister = "Member - " & rs.Fields("Debtor_Name") & " (" & TillData.Account_No & ")"
        frmSales.lblDebtor = "Member - " & rs.Fields("Debtor_Name") & " (" & TillData.Account_No & ")"
    Else
        frmSales.lblDebtor = ""
    End If
    rs.Close
    Panel_no = 0
    Select Case GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="TillSlip", Default:=1)
        Case 0
            frmSales.cmdSlip.Caption = "Slip off"
        Case 1
            frmSales.cmdSlip.Caption = "Slip on"
    End Select
    Select Case Branch_Type
        Case 8, 10
            cmdAppro.Caption = "PDT Upload"
        Case 11
            cmdAppro.Caption = "Fuel Pumps"
        Case 12
            cmdAppro.Caption = "Appro's"
        Case Else
            cmdAppro.Caption = "Appro's"
    End Select
    If Me.Height < 10000 And newBack.Visible = False Then
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.782
            Me.Controls(i).Left = Me.Controls(i).Left * 0.778
            Me.Controls(i).Height = Me.Controls(i).Height * 0.78
            Me.Controls(i).top = Me.Controls(i).top * 0.78
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
        newBack.Width = Screen.Width
        newBack.Height = Screen.Height
    End If

    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.6
    grdMain.ColWidth(2) = grdMain.Width * 0.25
    
    grdMain.ColWidth(14) = 200
    If frmSales.Tag = "1" Then
        frmSales.Tag = ""
        frmSales.lblKeyRegister.TextAlign = fmTextAlignLeft
    lblKeyRegister.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    If picSearch.Visible = False Then frmSales.KeyPreview = True
    If TillData.DocNo = 0 Then grdMain.Rows = 1
    picHoldFocus.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    frmSales.lblKeyRegister.TextAlign = fmTextAlignLeft
    lblKeyRegister.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    If picSearch.Visible = False Then frmSales.KeyPreview = True
    If TillData.DocNo = 0 Then grdMain.Rows = 1
    picHoldFocus.SetFocus
    cmdInput(15).Caption = "Void All"
    On Error GoTo 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Select Case KeyCode
            Case 83
                frmSales.Tag = "1"
                cmdInput_Click 18
        End Select
    End If
    If cmdErr.Visible = True Then
        frmSales.cmdErr.Visible = False
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
            grdMain.SetFocus
            DoEvents
        Case 40
            grdMain.SetFocus
            DoEvents
        Case 27
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
                frmSales.cmdErr.Caption = "Finalize the Sale before Logging Off"
                frmSales.cmdErr.Visible = True
                frmSales.errTimer.Enabled = True
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
            frmSales.KeyPreview = False
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
            grdMain.Rows = 1
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
                            Case "Reservations"
                                frmMain.cmdBar(0).Enabled = False
                                frmRes.Show
                            Case "Rooms"
                                frmMain.cmdBar(1).Enabled = False
                                frmRooms.Show
                            Case "Guests"
                                frmMain.cmdBar(2).Enabled = False
                                frmGuests.Show
                            Case "Checkin"
                                frmMain.cmdBar(3).Enabled = False
                                frmCheck.Show
                            Case "Stock"
                                frmMain.cmdBar(7).Enabled = False
                                frmProducts.Show
                            Case "Users"
                                frmMain.cmdBar(4).Enabled = False
                                frmUsers.Show
                            Case "Reports"
                                frmMain.cmdBar(5).Enabled = False
                                frmReports.Show
                            Case Else
                                frmDetails.Show
                        End Select
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
                    Me.Hide
                Case Else
                    Screen.MousePointer = 11
                    frmMain.cmdBar(6).Enabled = True
                    frmMain.Show
                    Select Case frmSales.Tag
                        Case "Reservations"
                            frmMain.cmdBar(0).Enabled = False
                            frmRes.Show
                        Case "Rooms"
                            frmMain.cmdBar(1).Enabled = False
                            frmRooms.Show
                        Case "Guests"
                            frmMain.cmdBar(2).Enabled = False
                            frmGuests.Show
                        Case "Checkin"
                            frmMain.cmdBar(3).Enabled = False
                            frmCheck.Show
                        Case "Stock"
                            frmMain.cmdBar(7).Enabled = False
                            frmProducts.Show
                        Case "Users"
                            frmMain.cmdBar(4).Enabled = False
                            frmUsers.Show
                        Case "Reports"
                            frmMain.cmdBar(5).Enabled = False
                            frmReports.Show
                        Case Else
                            frmDetails.Show
                    End Select
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
            Exit Sub
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
        Key_Function cmdKey(4).Caption
        
         Case 121
        Key_Function cmdKey(3).Caption
        
         Case 122
        Key_Function cmdKey(2).Caption
        
         Case 123
        Key_Function cmdKey(1).Caption
        
        
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
        Case 32
            Key_Function (Chr(KeyAscii))
        Case 43
            Key_Function "Price O/V"
        Case 42
            Key_Function "x"
        Case 46
            Key_Function "."
        Case 48 To 57
            Key_Function (Chr(KeyAscii))
        Case 65 To 90
            Key_Function (Chr(KeyAscii))
        Case 97 To 122
            KeyAscii = KeyAscii - 32
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
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
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
            grdMain.ShowCell grdMain.Row - 1, 0
    End Select
    picHoldFocus.SetFocus
    On Error GoTo 0
End Sub
Private Sub Form_Load()
    lblTender.Caption = "0.00"
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
    grdMain.Cell(flexcpForeColor, 0, 0, 0, 2) = &H80000012
    newBack.Width = Screen.Width
    newBack.Height = Screen.Height
End Sub

Private Sub grdFind_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    grdFind.ShowCell NewRow, 0
    rowfactor = 6865 / grdFind.Rows
    picKey.top = 2010 + (rowfactor * grdFind.Row)
    If NewRow = 1 Then picKey.top = 2010
    If NewRow = grdFind.Rows - 1 Then picKey.top = 8860
End Sub

Private Sub grdFind_DblClick()
    cmdArr(2).Tag = grdFind.TextMatrix(grdFind.Row, 0)
    If Panel_no = 0 Then frmSales.KeyPreview = True
    picSearch.Visible = False
    If InStr(KeyRegister, Chr$(215) & " *") <> 0 Then
        lblKeyRegister.Caption = Mid(lblKeyRegister.Caption, 1, InStr(lblKeyRegister.Caption, Chr$(215)) + 1)
    End If
    If InStr(KeyRegister, " (Return Item) ") <> 0 Then
        If InStr(KeyRegister, Chr$(215)) = 0 Then
            lblKeyRegister = " (Return Item) "
        End If
    End If
    KeyRegister = lblKeyRegister.Caption & cmdArr(2).Tag
    Key_Function "Plu"
    picHoldFocus.SetFocus
End Sub

Private Sub grdFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdArr(2).Tag = grdFind.TextMatrix(grdFind.Row, 0)
        If Panel_no = 0 Then frmSales.KeyPreview = True
        picSearch.Visible = False
        If InStr(KeyRegister, Chr$(215) & " *") <> 0 Then
            lblKeyRegister.Caption = Mid(lblKeyRegister.Caption, 1, InStr(lblKeyRegister.Caption, Chr$(215)) + 1)
        End If
        If InStr(KeyRegister, " (Return Item) ") <> 0 Then
            If InStr(KeyRegister, Chr$(215)) = 0 Then
                lblKeyRegister = " (Return Item) "
            End If
        End If
        KeyRegister = lblKeyRegister.Caption & cmdArr(2).Tag
        Key_Function "Plu"
        picHoldFocus.SetFocus
    End If
    If KeyCode = 27 Then
        If Panel_no = 0 Then frmSales.KeyPreview = True
        KeyRegister = ""
        picSearch.Visible = False
    End If
    If Shift = 2 And KeyCode = 75 Then
        If picSearch.Visible = True Then
            cmdKeyboard_Click
            Exit Sub
        End If
    End If
End Sub

Private Sub grdMain_Click()
    If Finalizing = True Then Exit Sub
    If GlobalMode = TillMode.StartMode Or GlobalMode = TillMode.NewMode Then
        grdMain.Rows = 1
        lblTender.Caption = "0.00"
        lblCash = ""
        lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    End If
    If grdMain.Cell(flexcpForeColor, grdMain.Row, 0, grdMain.Row, 2) = vbRed And TillData.DocNo <> 0 Then
       cmdErr.Caption = "This line was already Cleared"
        cmdErr.Visible = True
        errTimer.Enabled = True
        picHoldFocus.SetFocus
        Exit Sub
    End If
    If grdMain.TextMatrix(grdMain.Row, 8) = "Corr" And TillData.DocNo <> 0 Then
        cmdErr.Caption = "You cannot Void a Item Correct! Resell it"
        cmdErr.Visible = True
        errTimer.Enabled = True
        picHoldFocus.SetFocus
        Exit Sub
    End If
    If grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = &HC0C0FF And TillData.DocNo <> 0 Then
        cmdErr.Caption = "You cannot Void this Line! Re-sell it"
        cmdErr.Visible = True
        errTimer.Enabled = True
        picHoldFocus.SetFocus
        Exit Sub
    End If
    
    If KeyRegister = " (Void) " Then
        If grdMain.TextMatrix(grdMain.Row, 9) <> "0" Then
            If Trim(grdMain.TextMatrix(grdMain.Row, 0)) <> "" Then
                KeyRegister = " (Void) " & grdMain.TextMatrix(grdMain.Row, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(grdMain.Row, 2)) / Val(grdMain.TextMatrix(grdMain.Row, 0))) * 100 & " (Price O/V) " & grdMain.TextMatrix(grdMain.Row, 9)
                Key_Function "Plu"
            Else
                cmdErr.Caption = "You cannot Void this Line! Void the Item"
                cmdErr.Visible = True
                errTimer.Enabled = True
                picHoldFocus.SetFocus
                Exit Sub
            End If
        Else
            If Trim(grdMain.TextMatrix(grdMain.Row, 0)) <> "" Then
                KeyRegister = " (Void) " & grdMain.TextMatrix(grdMain.Row, 0) & " " & Chr(215) & " " & (Val(grdMain.TextMatrix(grdMain.Row, 2)) / Val(grdMain.TextMatrix(grdMain.Row, 0))) * 100 & " (Price O/V) " & grdMain.TextMatrix(grdMain.Row, 10)
                Key_Function "Dept"
            Else
                cmdErr.Caption = "You cannot Void this Line! Void the Item"
                cmdErr.Visible = True
                errTimer.Enabled = True
                picHoldFocus.SetFocus
                Exit Sub
            End If
        End If
        picHoldFocus.SetFocus
        Exit Sub
    End If
    cmdInput(15).Caption = "Void"
    If grdMain.Row > 0 And (GlobalMode = TillMode.Inputmode Or GlobalMode = TillMode.StartMode) Then
        If KeyRegister = "" And Trim(grdMain.TextMatrix(grdMain.Row, 0)) <> "" Then
            grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = &H87CDEB
            grdMain.HighLight = flexHighlightNever
            grdMain.Enabled = False
            voidTimer.Enabled = True
        End If
    End If
End Sub

Private Sub grdMain_GotFocus()
      grdMain.Tag = "1"
End Sub

Private Sub grdMain_LostFocus()
    grdMain.Tag = ""
End Sub

Private Sub scrolTimer_Timer()
    scrolTimer.Interval = 50
    Select Case scrolTimer.Tag
        Case "0"
            If grdFind.Row <> 1 Then
                grdFind.Row = grdFind.Row - 1
            End If
        Case "1"
            If grdFind.Row <> grdFind.Rows - 1 Then
                grdFind.Row = grdFind.Row + 1
            End If
    End Select
    grdFind.ShowCell grdFind.Row, 0
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
        Case &H87CDEB
            grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = &HC7E3F1
        Case &HC7E3F1
            grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = &H87CDEB
    End Select
End Sub
