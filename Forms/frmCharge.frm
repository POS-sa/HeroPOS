VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmCharge 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCharge.frx":0000
   ScaleHeight     =   9540
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTender 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6645
      Left            =   300
      ScaleHeight     =   6645
      ScaleWidth      =   3435
      TabIndex        =   43
      Top             =   2610
      Visible         =   0   'False
      Width           =   3435
      Begin BTNENHLib4.BtnEnh cmdTender 
         Height          =   1335
         Index           =   7
         Left            =   0
         TabIndex        =   44
         Top             =   6675
         Width           =   3375
         _Version        =   524298
         _ExtentX        =   5953
         _ExtentY        =   2355
         _StockProps     =   66
         Caption         =   "EFT"
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
         CornerFactor    =   15
         Surface         =   1
         BackColorContainer=   12632256
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":88FC
         textLT          =   "frmCharge.frx":8962
         textCT          =   "frmCharge.frx":897A
         textRT          =   "frmCharge.frx":8992
         textLM          =   "frmCharge.frx":89AA
         textRM          =   "frmCharge.frx":89C2
         textLB          =   "frmCharge.frx":89DA
         textCB          =   "frmCharge.frx":89F2
         textRB          =   "frmCharge.frx":8A0A
         colorBack       =   "frmCharge.frx":8A22
         colorIntern     =   "frmCharge.frx":8A4C
         colorMO         =   "frmCharge.frx":8A76
         colorFocus      =   "frmCharge.frx":8AA0
         colorDisabled   =   "frmCharge.frx":8ACA
         colorPressed    =   "frmCharge.frx":8AF4
         Style           =   2
         Orientation     =   1
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdTender 
         Height          =   1635
         Index           =   1
         Left            =   0
         TabIndex        =   47
         Top             =   1665
         Width           =   3435
         _Version        =   524298
         _ExtentX        =   6059
         _ExtentY        =   2884
         _StockProps     =   66
         Caption         =   "Card"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         CornerFactor    =   15
         Surface         =   1
         BackColorContainer=   1408168
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":8B1E
         textLT          =   "frmCharge.frx":8B86
         textCT          =   "frmCharge.frx":8B9E
         textRT          =   "frmCharge.frx":8BB6
         textLM          =   "frmCharge.frx":8BCE
         textRM          =   "frmCharge.frx":8BE6
         textLB          =   "frmCharge.frx":8BFE
         textCB          =   "frmCharge.frx":8C16
         textRB          =   "frmCharge.frx":8C2E
         colorBack       =   "frmCharge.frx":8C46
         colorIntern     =   "frmCharge.frx":8C70
         colorMO         =   "frmCharge.frx":8C9A
         colorFocus      =   "frmCharge.frx":8CC4
         colorDisabled   =   "frmCharge.frx":8CEE
         colorPressed    =   "frmCharge.frx":8D18
         Style           =   2
         Orientation     =   1
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdTender 
         Height          =   1665
         Index           =   0
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   3435
         _Version        =   524298
         _ExtentX        =   6059
         _ExtentY        =   2937
         _StockProps     =   66
         Caption         =   "Cash"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         CornerFactor    =   15
         Surface         =   1
         BackColorContainer=   1408168
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":8D42
         textLT          =   "frmCharge.frx":8DAA
         textCT          =   "frmCharge.frx":8DC2
         textRT          =   "frmCharge.frx":8DDA
         textLM          =   "frmCharge.frx":8DF2
         textRM          =   "frmCharge.frx":8E0A
         textLB          =   "frmCharge.frx":8E22
         textCB          =   "frmCharge.frx":8E3A
         textRB          =   "frmCharge.frx":8E52
         colorBack       =   "frmCharge.frx":8E6A
         colorIntern     =   "frmCharge.frx":8E94
         colorMO         =   "frmCharge.frx":8EBE
         colorFocus      =   "frmCharge.frx":8EE8
         colorDisabled   =   "frmCharge.frx":8F12
         colorPressed    =   "frmCharge.frx":8F3C
         Style           =   2
         Orientation     =   1
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdTender 
         Height          =   1635
         Index           =   4
         Left            =   0
         TabIndex        =   49
         Top             =   3300
         Width           =   3435
         _Version        =   524298
         _ExtentX        =   6059
         _ExtentY        =   2884
         _StockProps     =   66
         Caption         =   "Voucher"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         CornerFactor    =   15
         Surface         =   1
         BackColorContainer=   1408168
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":8F66
         textLT          =   "frmCharge.frx":8FD4
         textCT          =   "frmCharge.frx":8FEC
         textRT          =   "frmCharge.frx":9004
         textLM          =   "frmCharge.frx":901C
         textRM          =   "frmCharge.frx":9034
         textLB          =   "frmCharge.frx":904C
         textCB          =   "frmCharge.frx":9064
         textRB          =   "frmCharge.frx":907C
         colorBack       =   "frmCharge.frx":9094
         colorIntern     =   "frmCharge.frx":90BE
         colorMO         =   "frmCharge.frx":90E8
         colorFocus      =   "frmCharge.frx":9112
         colorDisabled   =   "frmCharge.frx":913C
         colorPressed    =   "frmCharge.frx":9166
         Style           =   2
         Orientation     =   1
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdTender 
         Height          =   1695
         Index           =   2
         Left            =   0
         TabIndex        =   50
         Top             =   4935
         Width           =   3435
         _Version        =   524298
         _ExtentX        =   6059
         _ExtentY        =   2990
         _StockProps     =   66
         Caption         =   "EFT"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         CornerFactor    =   15
         Surface         =   1
         BackColorContainer=   1408168
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":9190
         textLT          =   "frmCharge.frx":91F6
         textCT          =   "frmCharge.frx":920E
         textRT          =   "frmCharge.frx":9226
         textLM          =   "frmCharge.frx":923E
         textRM          =   "frmCharge.frx":9256
         textLB          =   "frmCharge.frx":926E
         textCB          =   "frmCharge.frx":9286
         textRB          =   "frmCharge.frx":929E
         colorBack       =   "frmCharge.frx":92B6
         colorIntern     =   "frmCharge.frx":92E0
         colorMO         =   "frmCharge.frx":930A
         colorFocus      =   "frmCharge.frx":9334
         colorDisabled   =   "frmCharge.frx":935E
         colorPressed    =   "frmCharge.frx":9388
         Style           =   2
         Orientation     =   1
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
   End
   Begin VB.Timer errTimer 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picReceive 
      Height          =   6645
      Left            =   300
      ScaleHeight     =   6585
      ScaleWidth      =   8685
      TabIndex        =   24
      Top             =   2610
      Visible         =   0   'False
      Width           =   8745
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1395
         Index           =   0
         Left            =   3420
         TabIndex        =   26
         Top             =   0
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   2461
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
         textCaption     =   "frmCharge.frx":93B2
         textLT          =   "frmCharge.frx":9414
         textCT          =   "frmCharge.frx":942C
         textRT          =   "frmCharge.frx":9444
         textLM          =   "frmCharge.frx":945C
         textRM          =   "frmCharge.frx":9474
         textLB          =   "frmCharge.frx":948C
         textCB          =   "frmCharge.frx":94A4
         textRB          =   "frmCharge.frx":94BC
         colorBack       =   "frmCharge.frx":94D4
         colorIntern     =   "frmCharge.frx":94FE
         colorMO         =   "frmCharge.frx":9528
         colorFocus      =   "frmCharge.frx":9552
         colorDisabled   =   "frmCharge.frx":957C
         colorPressed    =   "frmCharge.frx":95A6
         Orientation     =   5
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1365
         Index           =   9
         Left            =   3420
         TabIndex        =   27
         Top             =   4125
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   2408
         _StockProps     =   66
         Caption         =   "CL"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         CornerFactor    =   15
         BackColorContainer=   12632256
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":95D0
         textLT          =   "frmCharge.frx":9634
         textCT          =   "frmCharge.frx":964C
         textRT          =   "frmCharge.frx":9664
         textLM          =   "frmCharge.frx":967C
         textRM          =   "frmCharge.frx":9694
         textLB          =   "frmCharge.frx":96AC
         textCB          =   "frmCharge.frx":96C4
         textRB          =   "frmCharge.frx":96DC
         colorBack       =   "frmCharge.frx":96F4
         colorIntern     =   "frmCharge.frx":971E
         colorMO         =   "frmCharge.frx":9748
         colorFocus      =   "frmCharge.frx":9772
         colorDisabled   =   "frmCharge.frx":979C
         colorPressed    =   "frmCharge.frx":97C6
         Orientation     =   8
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1395
         Index           =   6
         Left            =   3420
         TabIndex        =   28
         Top             =   2730
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   2461
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
         textCaption     =   "frmCharge.frx":97F0
         textLT          =   "frmCharge.frx":9852
         textCT          =   "frmCharge.frx":986A
         textRT          =   "frmCharge.frx":9882
         textLM          =   "frmCharge.frx":989A
         textRM          =   "frmCharge.frx":98B2
         textLB          =   "frmCharge.frx":98CA
         textCB          =   "frmCharge.frx":98E2
         textRB          =   "frmCharge.frx":98FA
         colorBack       =   "frmCharge.frx":9912
         colorIntern     =   "frmCharge.frx":993C
         colorMO         =   "frmCharge.frx":9966
         colorFocus      =   "frmCharge.frx":9990
         colorDisabled   =   "frmCharge.frx":99BA
         colorPressed    =   "frmCharge.frx":99E4
         Orientation     =   5
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1335
         Index           =   3
         Left            =   3420
         TabIndex        =   29
         Top             =   1395
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   2355
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
         textCaption     =   "frmCharge.frx":9A0E
         textLT          =   "frmCharge.frx":9A70
         textCT          =   "frmCharge.frx":9A88
         textRT          =   "frmCharge.frx":9AA0
         textLM          =   "frmCharge.frx":9AB8
         textRM          =   "frmCharge.frx":9AD0
         textLB          =   "frmCharge.frx":9AE8
         textCB          =   "frmCharge.frx":9B00
         textRB          =   "frmCharge.frx":9B18
         colorBack       =   "frmCharge.frx":9B30
         colorIntern     =   "frmCharge.frx":9B5A
         colorMO         =   "frmCharge.frx":9B84
         colorFocus      =   "frmCharge.frx":9BAE
         colorDisabled   =   "frmCharge.frx":9BD8
         colorPressed    =   "frmCharge.frx":9C02
         Orientation     =   5
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1365
         Index           =   10
         Left            =   5175
         TabIndex        =   30
         Top             =   4125
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   2408
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
         textCaption     =   "frmCharge.frx":9C2C
         textLT          =   "frmCharge.frx":9C8E
         textCT          =   "frmCharge.frx":9CA6
         textRT          =   "frmCharge.frx":9CBE
         textLM          =   "frmCharge.frx":9CD6
         textRM          =   "frmCharge.frx":9CEE
         textLB          =   "frmCharge.frx":9D06
         textCB          =   "frmCharge.frx":9D1E
         textRB          =   "frmCharge.frx":9D36
         colorBack       =   "frmCharge.frx":9D4E
         colorIntern     =   "frmCharge.frx":9D78
         colorMO         =   "frmCharge.frx":9DA2
         colorFocus      =   "frmCharge.frx":9DCC
         colorDisabled   =   "frmCharge.frx":9DF6
         colorPressed    =   "frmCharge.frx":9E20
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1395
         Index           =   7
         Left            =   5175
         TabIndex        =   31
         Top             =   2730
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   2461
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
         textCaption     =   "frmCharge.frx":9E4A
         textLT          =   "frmCharge.frx":9EAC
         textCT          =   "frmCharge.frx":9EC4
         textRT          =   "frmCharge.frx":9EDC
         textLM          =   "frmCharge.frx":9EF4
         textRM          =   "frmCharge.frx":9F0C
         textLB          =   "frmCharge.frx":9F24
         textCB          =   "frmCharge.frx":9F3C
         textRB          =   "frmCharge.frx":9F54
         colorBack       =   "frmCharge.frx":9F6C
         colorIntern     =   "frmCharge.frx":9F96
         colorMO         =   "frmCharge.frx":9FC0
         colorFocus      =   "frmCharge.frx":9FEA
         colorDisabled   =   "frmCharge.frx":A014
         colorPressed    =   "frmCharge.frx":A03E
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1335
         Index           =   4
         Left            =   5175
         TabIndex        =   32
         Top             =   1395
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   2355
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
         textCaption     =   "frmCharge.frx":A068
         textLT          =   "frmCharge.frx":A0CA
         textCT          =   "frmCharge.frx":A0E2
         textRT          =   "frmCharge.frx":A0FA
         textLM          =   "frmCharge.frx":A112
         textRM          =   "frmCharge.frx":A12A
         textLB          =   "frmCharge.frx":A142
         textCB          =   "frmCharge.frx":A15A
         textRB          =   "frmCharge.frx":A172
         colorBack       =   "frmCharge.frx":A18A
         colorIntern     =   "frmCharge.frx":A1B4
         colorMO         =   "frmCharge.frx":A1DE
         colorFocus      =   "frmCharge.frx":A208
         colorDisabled   =   "frmCharge.frx":A232
         colorPressed    =   "frmCharge.frx":A25C
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1365
         Index           =   11
         Left            =   6930
         TabIndex        =   33
         Top             =   4125
         Width           =   1785
         _Version        =   524298
         _ExtentX        =   3149
         _ExtentY        =   2408
         _StockProps     =   66
         Caption         =   "OK"
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
         textCaption     =   "frmCharge.frx":A286
         textLT          =   "frmCharge.frx":A2EA
         textCT          =   "frmCharge.frx":A302
         textRT          =   "frmCharge.frx":A31A
         textLM          =   "frmCharge.frx":A332
         textRM          =   "frmCharge.frx":A34A
         textLB          =   "frmCharge.frx":A362
         textCB          =   "frmCharge.frx":A37A
         textRB          =   "frmCharge.frx":A392
         colorBack       =   "frmCharge.frx":A3AA
         colorIntern     =   "frmCharge.frx":A3D4
         colorMO         =   "frmCharge.frx":A3FE
         colorFocus      =   "frmCharge.frx":A428
         colorDisabled   =   "frmCharge.frx":A452
         colorPressed    =   "frmCharge.frx":A47C
         Orientation     =   7
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1395
         Index           =   8
         Left            =   6930
         TabIndex        =   34
         Top             =   2730
         Width           =   1785
         _Version        =   524298
         _ExtentX        =   3149
         _ExtentY        =   2461
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
         textCaption     =   "frmCharge.frx":A4A6
         textLT          =   "frmCharge.frx":A508
         textCT          =   "frmCharge.frx":A520
         textRT          =   "frmCharge.frx":A538
         textLM          =   "frmCharge.frx":A550
         textRM          =   "frmCharge.frx":A568
         textLB          =   "frmCharge.frx":A580
         textCB          =   "frmCharge.frx":A598
         textRB          =   "frmCharge.frx":A5B0
         colorBack       =   "frmCharge.frx":A5C8
         colorIntern     =   "frmCharge.frx":A5F2
         colorMO         =   "frmCharge.frx":A61C
         colorFocus      =   "frmCharge.frx":A646
         colorDisabled   =   "frmCharge.frx":A670
         colorPressed    =   "frmCharge.frx":A69A
         Orientation     =   6
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1335
         Index           =   5
         Left            =   6930
         TabIndex        =   35
         Top             =   1395
         Width           =   1785
         _Version        =   524298
         _ExtentX        =   3149
         _ExtentY        =   2355
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
         textCaption     =   "frmCharge.frx":A6C4
         textLT          =   "frmCharge.frx":A726
         textCT          =   "frmCharge.frx":A73E
         textRT          =   "frmCharge.frx":A756
         textLM          =   "frmCharge.frx":A76E
         textRM          =   "frmCharge.frx":A786
         textLB          =   "frmCharge.frx":A79E
         textCB          =   "frmCharge.frx":A7B6
         textRB          =   "frmCharge.frx":A7CE
         colorBack       =   "frmCharge.frx":A7E6
         colorIntern     =   "frmCharge.frx":A810
         colorMO         =   "frmCharge.frx":A83A
         colorFocus      =   "frmCharge.frx":A864
         colorDisabled   =   "frmCharge.frx":A88E
         colorPressed    =   "frmCharge.frx":A8B8
         Orientation     =   6
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1395
         Index           =   1
         Left            =   5175
         TabIndex        =   36
         Top             =   0
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   2461
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
         textCaption     =   "frmCharge.frx":A8E2
         textLT          =   "frmCharge.frx":A944
         textCT          =   "frmCharge.frx":A95C
         textRT          =   "frmCharge.frx":A974
         textLM          =   "frmCharge.frx":A98C
         textRM          =   "frmCharge.frx":A9A4
         textLB          =   "frmCharge.frx":A9BC
         textCB          =   "frmCharge.frx":A9D4
         textRB          =   "frmCharge.frx":A9EC
         colorBack       =   "frmCharge.frx":AA04
         colorIntern     =   "frmCharge.frx":AA2E
         colorMO         =   "frmCharge.frx":AA58
         colorFocus      =   "frmCharge.frx":AA82
         colorDisabled   =   "frmCharge.frx":AAAC
         colorPressed    =   "frmCharge.frx":AAD6
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1395
         Index           =   2
         Left            =   6930
         TabIndex        =   37
         Top             =   0
         Width           =   1785
         _Version        =   524298
         _ExtentX        =   3149
         _ExtentY        =   2461
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
         textCaption     =   "frmCharge.frx":AB00
         textLT          =   "frmCharge.frx":AB62
         textCT          =   "frmCharge.frx":AB7A
         textRT          =   "frmCharge.frx":AB92
         textLM          =   "frmCharge.frx":ABAA
         textRM          =   "frmCharge.frx":ABC2
         textLB          =   "frmCharge.frx":ABDA
         textCB          =   "frmCharge.frx":ABF2
         textRB          =   "frmCharge.frx":AC0A
         colorBack       =   "frmCharge.frx":AC22
         colorIntern     =   "frmCharge.frx":AC4C
         colorMO         =   "frmCharge.frx":AC76
         colorFocus      =   "frmCharge.frx":ACA0
         colorDisabled   =   "frmCharge.frx":ACCA
         colorPressed    =   "frmCharge.frx":ACF4
         Orientation     =   6
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdPrint 
         Height          =   1065
         Index           =   0
         Left            =   3420
         TabIndex        =   38
         Top             =   5520
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   1879
         _StockProps     =   66
         Caption         =   "Slip"
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
         Shape           =   1
         CornerFactor    =   15
         BackColorContainer=   12632256
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":AD1E
         textLT          =   "frmCharge.frx":AD86
         textCT          =   "frmCharge.frx":AD9E
         textRT          =   "frmCharge.frx":ADB6
         textLM          =   "frmCharge.frx":ADCE
         textRM          =   "frmCharge.frx":ADE6
         textLB          =   "frmCharge.frx":ADFE
         textCB          =   "frmCharge.frx":AE16
         textRB          =   "frmCharge.frx":AE2E
         colorBack       =   "frmCharge.frx":AE46
         colorIntern     =   "frmCharge.frx":AE70
         colorMO         =   "frmCharge.frx":AE9A
         colorFocus      =   "frmCharge.frx":AEC4
         colorDisabled   =   "frmCharge.frx":AEEE
         colorPressed    =   "frmCharge.frx":AF18
         Style           =   2
         Orientation     =   8
         HollowFrame     =   -1  'True
         LightDirection  =   7
         Value           =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdKey 
         Height          =   1065
         Index           =   12
         Left            =   5175
         TabIndex        =   39
         Top             =   5520
         Width           =   1755
         _Version        =   524298
         _ExtentX        =   3096
         _ExtentY        =   1879
         _StockProps     =   66
         Caption         =   "."
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   36
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
         BackColorContainer=   16777215
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":AF42
         textLT          =   "frmCharge.frx":AFA4
         textCT          =   "frmCharge.frx":AFBC
         textRT          =   "frmCharge.frx":AFD4
         textLM          =   "frmCharge.frx":AFEC
         textRM          =   "frmCharge.frx":B004
         textLB          =   "frmCharge.frx":B01C
         textCB          =   "frmCharge.frx":B034
         textRB          =   "frmCharge.frx":B04C
         colorBack       =   "frmCharge.frx":B064
         colorIntern     =   "frmCharge.frx":B08E
         colorMO         =   "frmCharge.frx":B0B8
         colorFocus      =   "frmCharge.frx":B0E2
         colorDisabled   =   "frmCharge.frx":B10C
         colorPressed    =   "frmCharge.frx":B136
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdPrint 
         Height          =   1065
         Index           =   1
         Left            =   6930
         TabIndex        =   40
         Top             =   5520
         Width           =   1785
         _Version        =   524298
         _ExtentX        =   3149
         _ExtentY        =   1879
         _StockProps     =   66
         Caption         =   "A4"
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
         Shape           =   1
         CornerFactor    =   15
         BackColorContainer=   12632256
         SpecialEffect   =   1
         LogPixels       =   96
         Clickable       =   0   'False
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmCharge.frx":B160
         textLT          =   "frmCharge.frx":B1C4
         textCT          =   "frmCharge.frx":B1DC
         textRT          =   "frmCharge.frx":B1F4
         textLM          =   "frmCharge.frx":B20C
         textRM          =   "frmCharge.frx":B224
         textLB          =   "frmCharge.frx":B23C
         textCB          =   "frmCharge.frx":B254
         textRB          =   "frmCharge.frx":B26C
         colorBack       =   "frmCharge.frx":B284
         colorIntern     =   "frmCharge.frx":B2AE
         colorMO         =   "frmCharge.frx":B2D8
         colorFocus      =   "frmCharge.frx":B302
         colorDisabled   =   "frmCharge.frx":B32C
         colorPressed    =   "frmCharge.frx":B356
         Style           =   2
         Orientation     =   7
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin VSFlex8Ctl.VSFlexGrid grdMain 
         Height          =   6000
         Left            =   30
         TabIndex        =   42
         Top             =   570
         Width           =   3345
         _cx             =   5900
         _cy             =   10583
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16744576
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16381166
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSForms.Label lblTender 
         Height          =   435
         Left            =   1860
         TabIndex        =   25
         Top             =   90
         Width           =   1335
         ForeColor       =   7555868
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "0.00"
         Size            =   "2355;767"
         FontName        =   "Arial Narrow"
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblCash 
         Height          =   435
         Left            =   180
         TabIndex        =   41
         Top             =   90
         Width           =   1245
         ForeColor       =   7555868
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Balance"
         Size            =   "2196;767"
         FontName        =   "Arial Narrow"
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image Image1 
         Height          =   525
         Left            =   30
         Top             =   30
         Width           =   3360
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "5918;926"
      End
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   870
      Index           =   0
      Left            =   7920
      TabIndex        =   1
      Top             =   270
      Width           =   1125
      _ExtentX        =   1984
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
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   885
      Left            =   300
      TabIndex        =   2
      Top             =   330
      Visible         =   0   'False
      Width           =   7485
      _Version        =   524298
      _ExtentX        =   13203
      _ExtentY        =   1561
      _StockProps     =   66
      Caption         =   "Invalid Key Pressed"
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
      textCaption     =   "frmCharge.frx":B380
      textLT          =   "frmCharge.frx":B406
      textCT          =   "frmCharge.frx":B41E
      textRT          =   "frmCharge.frx":B436
      textLM          =   "frmCharge.frx":B44E
      textRM          =   "frmCharge.frx":B466
      textLB          =   "frmCharge.frx":B47E
      textCB          =   "frmCharge.frx":B496
      textRB          =   "frmCharge.frx":B4AE
      colorBack       =   "frmCharge.frx":B4C6
      colorIntern     =   "frmCharge.frx":B4F0
      colorMO         =   "frmCharge.frx":B51A
      colorFocus      =   "frmCharge.frx":B544
      colorDisabled   =   "frmCharge.frx":B56E
      colorPressed    =   "frmCharge.frx":B598
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1185
      Index           =   0
      Left            =   300
      TabIndex        =   4
      Top             =   1290
      Width           =   1395
      _Version        =   524298
      _ExtentX        =   2461
      _ExtentY        =   2090
      _StockProps     =   66
      Caption         =   "Debtor"
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":B5C2
      textLT          =   "frmCharge.frx":B62E
      textCT          =   "frmCharge.frx":B646
      textRT          =   "frmCharge.frx":B65E
      textLM          =   "frmCharge.frx":B676
      textRM          =   "frmCharge.frx":B68E
      textLB          =   "frmCharge.frx":B6A6
      textCB          =   "frmCharge.frx":B6BE
      textRB          =   "frmCharge.frx":B6D6
      colorBack       =   "frmCharge.frx":B6EE
      colorIntern     =   "frmCharge.frx":B718
      colorMO         =   "frmCharge.frx":B742
      colorFocus      =   "frmCharge.frx":B76C
      colorDisabled   =   "frmCharge.frx":B796
      colorPressed    =   "frmCharge.frx":B7C0
      Style           =   2
      Orientation     =   1
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1185
      Index           =   1
      Left            =   1710
      TabIndex        =   5
      Top             =   1290
      Width           =   1425
      _Version        =   524298
      _ExtentX        =   2514
      _ExtentY        =   2090
      _StockProps     =   66
      Caption         =   "Staff"
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":B7EA
      textLT          =   "frmCharge.frx":B854
      textCT          =   "frmCharge.frx":B86C
      textRT          =   "frmCharge.frx":B884
      textLM          =   "frmCharge.frx":B89C
      textRM          =   "frmCharge.frx":B8B4
      textLB          =   "frmCharge.frx":B8CC
      textCB          =   "frmCharge.frx":B8E4
      textRB          =   "frmCharge.frx":B8FC
      colorBack       =   "frmCharge.frx":B914
      colorIntern     =   "frmCharge.frx":B93E
      colorMO         =   "frmCharge.frx":B968
      colorFocus      =   "frmCharge.frx":B992
      colorDisabled   =   "frmCharge.frx":B9BC
      colorPressed    =   "frmCharge.frx":B9E6
      Style           =   2
      Orientation     =   1
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1185
      Index           =   2
      Left            =   3150
      TabIndex        =   6
      Top             =   1290
      Width           =   1665
      _Version        =   524298
      _ExtentX        =   2937
      _ExtentY        =   2090
      _StockProps     =   66
      Caption         =   "Management"
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":BA10
      textLT          =   "frmCharge.frx":BA84
      textCT          =   "frmCharge.frx":BA9C
      textRT          =   "frmCharge.frx":BAB4
      textLM          =   "frmCharge.frx":BACC
      textRM          =   "frmCharge.frx":BAE4
      textLB          =   "frmCharge.frx":BAFC
      textCB          =   "frmCharge.frx":BB14
      textRB          =   "frmCharge.frx":BB2C
      colorBack       =   "frmCharge.frx":BB44
      colorIntern     =   "frmCharge.frx":BB6E
      colorMO         =   "frmCharge.frx":BB98
      colorFocus      =   "frmCharge.frx":BBC2
      colorDisabled   =   "frmCharge.frx":BBEC
      colorPressed    =   "frmCharge.frx":BC16
      Style           =   2
      Orientation     =   1
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1185
      Index           =   3
      Left            =   4800
      TabIndex        =   7
      Top             =   1290
      Width           =   1425
      _Version        =   524298
      _ExtentX        =   2514
      _ExtentY        =   2090
      _StockProps     =   66
      Caption         =   "Room"
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":BC40
      textLT          =   "frmCharge.frx":BCA8
      textCT          =   "frmCharge.frx":BCC0
      textRT          =   "frmCharge.frx":BCD8
      textLM          =   "frmCharge.frx":BCF0
      textRM          =   "frmCharge.frx":BD08
      textLB          =   "frmCharge.frx":BD20
      textCB          =   "frmCharge.frx":BD38
      textRB          =   "frmCharge.frx":BD50
      colorBack       =   "frmCharge.frx":BD68
      colorIntern     =   "frmCharge.frx":BD92
      colorMO         =   "frmCharge.frx":BDBC
      colorFocus      =   "frmCharge.frx":BDE6
      colorDisabled   =   "frmCharge.frx":BE10
      colorPressed    =   "frmCharge.frx":BE3A
      Style           =   2
      Orientation     =   3
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   3
      Left            =   300
      TabIndex        =   8
      Top             =   3945
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":BE64
      textLT          =   "frmCharge.frx":BE7C
      textCT          =   "frmCharge.frx":BE94
      textRT          =   "frmCharge.frx":BEAC
      textLM          =   "frmCharge.frx":BEC4
      textRM          =   "frmCharge.frx":BEDC
      textLB          =   "frmCharge.frx":BEF4
      textCB          =   "frmCharge.frx":BF0C
      textRB          =   "frmCharge.frx":BF24
      colorBack       =   "frmCharge.frx":BF3C
      colorIntern     =   "frmCharge.frx":BF66
      colorMO         =   "frmCharge.frx":BF90
      colorFocus      =   "frmCharge.frx":BFBA
      colorDisabled   =   "frmCharge.frx":BFE4
      colorPressed    =   "frmCharge.frx":C00E
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   5
      Left            =   6150
      TabIndex        =   9
      Top             =   3945
      Visible         =   0   'False
      Width           =   2895
      _Version        =   524298
      _ExtentX        =   5106
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":C038
      textLT          =   "frmCharge.frx":C050
      textCT          =   "frmCharge.frx":C068
      textRT          =   "frmCharge.frx":C080
      textLM          =   "frmCharge.frx":C098
      textRM          =   "frmCharge.frx":C0B0
      textLB          =   "frmCharge.frx":C0C8
      textCB          =   "frmCharge.frx":C0E0
      textRB          =   "frmCharge.frx":C0F8
      colorBack       =   "frmCharge.frx":C110
      colorIntern     =   "frmCharge.frx":C13A
      colorMO         =   "frmCharge.frx":C164
      colorFocus      =   "frmCharge.frx":C18E
      colorDisabled   =   "frmCharge.frx":C1B8
      colorPressed    =   "frmCharge.frx":C1E2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   6
      Left            =   300
      TabIndex        =   10
      Top             =   5250
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":C20C
      textLT          =   "frmCharge.frx":C224
      textCT          =   "frmCharge.frx":C23C
      textRT          =   "frmCharge.frx":C254
      textLM          =   "frmCharge.frx":C26C
      textRM          =   "frmCharge.frx":C284
      textLB          =   "frmCharge.frx":C29C
      textCB          =   "frmCharge.frx":C2B4
      textRB          =   "frmCharge.frx":C2CC
      colorBack       =   "frmCharge.frx":C2E4
      colorIntern     =   "frmCharge.frx":C30E
      colorMO         =   "frmCharge.frx":C338
      colorFocus      =   "frmCharge.frx":C362
      colorDisabled   =   "frmCharge.frx":C38C
      colorPressed    =   "frmCharge.frx":C3B6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   8
      Left            =   6150
      TabIndex        =   11
      Top             =   5250
      Visible         =   0   'False
      Width           =   2895
      _Version        =   524298
      _ExtentX        =   5106
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":C3E0
      textLT          =   "frmCharge.frx":C3F8
      textCT          =   "frmCharge.frx":C410
      textRT          =   "frmCharge.frx":C428
      textLM          =   "frmCharge.frx":C440
      textRM          =   "frmCharge.frx":C458
      textLB          =   "frmCharge.frx":C470
      textCB          =   "frmCharge.frx":C488
      textRB          =   "frmCharge.frx":C4A0
      colorBack       =   "frmCharge.frx":C4B8
      colorIntern     =   "frmCharge.frx":C4E2
      colorMO         =   "frmCharge.frx":C50C
      colorFocus      =   "frmCharge.frx":C536
      colorDisabled   =   "frmCharge.frx":C560
      colorPressed    =   "frmCharge.frx":C58A
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   9
      Left            =   300
      TabIndex        =   12
      Top             =   6555
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":C5B4
      textLT          =   "frmCharge.frx":C5CC
      textCT          =   "frmCharge.frx":C5E4
      textRT          =   "frmCharge.frx":C5FC
      textLM          =   "frmCharge.frx":C614
      textRM          =   "frmCharge.frx":C62C
      textLB          =   "frmCharge.frx":C644
      textCB          =   "frmCharge.frx":C65C
      textRB          =   "frmCharge.frx":C674
      colorBack       =   "frmCharge.frx":C68C
      colorIntern     =   "frmCharge.frx":C6B6
      colorMO         =   "frmCharge.frx":C6E0
      colorFocus      =   "frmCharge.frx":C70A
      colorDisabled   =   "frmCharge.frx":C734
      colorPressed    =   "frmCharge.frx":C75E
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   11
      Left            =   6150
      TabIndex        =   13
      Top             =   6555
      Visible         =   0   'False
      Width           =   2895
      _Version        =   524298
      _ExtentX        =   5106
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":C788
      textLT          =   "frmCharge.frx":C7A0
      textCT          =   "frmCharge.frx":C7B8
      textRT          =   "frmCharge.frx":C7D0
      textLM          =   "frmCharge.frx":C7E8
      textRM          =   "frmCharge.frx":C800
      textLB          =   "frmCharge.frx":C818
      textCB          =   "frmCharge.frx":C830
      textRB          =   "frmCharge.frx":C848
      colorBack       =   "frmCharge.frx":C860
      colorIntern     =   "frmCharge.frx":C88A
      colorMO         =   "frmCharge.frx":C8B4
      colorFocus      =   "frmCharge.frx":C8DE
      colorDisabled   =   "frmCharge.frx":C908
      colorPressed    =   "frmCharge.frx":C932
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   2
      Left            =   6150
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   2895
      _Version        =   524298
      _ExtentX        =   5106
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":C95C
      textLT          =   "frmCharge.frx":C974
      textCT          =   "frmCharge.frx":C98C
      textRT          =   "frmCharge.frx":C9A4
      textLM          =   "frmCharge.frx":C9BC
      textRM          =   "frmCharge.frx":C9D4
      textLB          =   "frmCharge.frx":C9EC
      textCB          =   "frmCharge.frx":CA04
      textRB          =   "frmCharge.frx":CA1C
      colorBack       =   "frmCharge.frx":CA34
      colorIntern     =   "frmCharge.frx":CA5E
      colorMO         =   "frmCharge.frx":CA88
      colorFocus      =   "frmCharge.frx":CAB2
      colorDisabled   =   "frmCharge.frx":CADC
      colorPressed    =   "frmCharge.frx":CB06
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   1
      Left            =   3225
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":CB30
      textLT          =   "frmCharge.frx":CB48
      textCT          =   "frmCharge.frx":CB60
      textRT          =   "frmCharge.frx":CB78
      textLM          =   "frmCharge.frx":CB90
      textRM          =   "frmCharge.frx":CBA8
      textLB          =   "frmCharge.frx":CBC0
      textCB          =   "frmCharge.frx":CBD8
      textRB          =   "frmCharge.frx":CBF0
      colorBack       =   "frmCharge.frx":CC08
      colorIntern     =   "frmCharge.frx":CC32
      colorMO         =   "frmCharge.frx":CC5C
      colorFocus      =   "frmCharge.frx":CC86
      colorDisabled   =   "frmCharge.frx":CCB0
      colorPressed    =   "frmCharge.frx":CCDA
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   4
      Left            =   3225
      TabIndex        =   16
      Top             =   3945
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":CD04
      textLT          =   "frmCharge.frx":CD1C
      textCT          =   "frmCharge.frx":CD34
      textRT          =   "frmCharge.frx":CD4C
      textLM          =   "frmCharge.frx":CD64
      textRM          =   "frmCharge.frx":CD7C
      textLB          =   "frmCharge.frx":CD94
      textCB          =   "frmCharge.frx":CDAC
      textRB          =   "frmCharge.frx":CDC4
      colorBack       =   "frmCharge.frx":CDDC
      colorIntern     =   "frmCharge.frx":CE06
      colorMO         =   "frmCharge.frx":CE30
      colorFocus      =   "frmCharge.frx":CE5A
      colorDisabled   =   "frmCharge.frx":CE84
      colorPressed    =   "frmCharge.frx":CEAE
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   7
      Left            =   3225
      TabIndex        =   17
      Top             =   5250
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":CED8
      textLT          =   "frmCharge.frx":CEF0
      textCT          =   "frmCharge.frx":CF08
      textRT          =   "frmCharge.frx":CF20
      textLM          =   "frmCharge.frx":CF38
      textRM          =   "frmCharge.frx":CF50
      textLB          =   "frmCharge.frx":CF68
      textCB          =   "frmCharge.frx":CF80
      textRB          =   "frmCharge.frx":CF98
      colorBack       =   "frmCharge.frx":CFB0
      colorIntern     =   "frmCharge.frx":CFDA
      colorMO         =   "frmCharge.frx":D004
      colorFocus      =   "frmCharge.frx":D02E
      colorDisabled   =   "frmCharge.frx":D058
      colorPressed    =   "frmCharge.frx":D082
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   10
      Left            =   3225
      TabIndex        =   18
      Top             =   6555
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":D0AC
      textLT          =   "frmCharge.frx":D0C4
      textCT          =   "frmCharge.frx":D0DC
      textRT          =   "frmCharge.frx":D0F4
      textLM          =   "frmCharge.frx":D10C
      textRM          =   "frmCharge.frx":D124
      textLB          =   "frmCharge.frx":D13C
      textCB          =   "frmCharge.frx":D154
      textRB          =   "frmCharge.frx":D16C
      colorBack       =   "frmCharge.frx":D184
      colorIntern     =   "frmCharge.frx":D1AE
      colorMO         =   "frmCharge.frx":D1D8
      colorFocus      =   "frmCharge.frx":D202
      colorDisabled   =   "frmCharge.frx":D22C
      colorPressed    =   "frmCharge.frx":D256
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1305
      Index           =   0
      Left            =   300
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":D280
      textLT          =   "frmCharge.frx":D298
      textCT          =   "frmCharge.frx":D2B0
      textRT          =   "frmCharge.frx":D2C8
      textLM          =   "frmCharge.frx":D2E0
      textRM          =   "frmCharge.frx":D2F8
      textLB          =   "frmCharge.frx":D310
      textCB          =   "frmCharge.frx":D328
      textRB          =   "frmCharge.frx":D340
      colorBack       =   "frmCharge.frx":D358
      colorIntern     =   "frmCharge.frx":D382
      colorMO         =   "frmCharge.frx":D3AC
      colorFocus      =   "frmCharge.frx":D3D6
      colorDisabled   =   "frmCharge.frx":D400
      colorPressed    =   "frmCharge.frx":D42A
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1365
      Index           =   12
      Left            =   300
      TabIndex        =   20
      Top             =   7860
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
      _ExtentY        =   2408
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":D454
      textLT          =   "frmCharge.frx":D46C
      textCT          =   "frmCharge.frx":D484
      textRT          =   "frmCharge.frx":D49C
      textLM          =   "frmCharge.frx":D4B4
      textRM          =   "frmCharge.frx":D4CC
      textLB          =   "frmCharge.frx":D4E4
      textCB          =   "frmCharge.frx":D4FC
      textRB          =   "frmCharge.frx":D514
      colorBack       =   "frmCharge.frx":D52C
      colorIntern     =   "frmCharge.frx":D556
      colorMO         =   "frmCharge.frx":D580
      colorFocus      =   "frmCharge.frx":D5AA
      colorDisabled   =   "frmCharge.frx":D5D4
      colorPressed    =   "frmCharge.frx":D5FE
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1365
      Index           =   13
      Left            =   3225
      TabIndex        =   21
      Top             =   7860
      Visible         =   0   'False
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
      _ExtentY        =   2408
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":D628
      textLT          =   "frmCharge.frx":D640
      textCT          =   "frmCharge.frx":D658
      textRT          =   "frmCharge.frx":D670
      textLM          =   "frmCharge.frx":D688
      textRM          =   "frmCharge.frx":D6A0
      textLB          =   "frmCharge.frx":D6B8
      textCB          =   "frmCharge.frx":D6D0
      textRB          =   "frmCharge.frx":D6E8
      colorBack       =   "frmCharge.frx":D700
      colorIntern     =   "frmCharge.frx":D72A
      colorMO         =   "frmCharge.frx":D754
      colorFocus      =   "frmCharge.frx":D77E
      colorDisabled   =   "frmCharge.frx":D7A8
      colorPressed    =   "frmCharge.frx":D7D2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdServers 
      Height          =   1365
      Index           =   14
      Left            =   6150
      TabIndex        =   22
      Top             =   7860
      Visible         =   0   'False
      Width           =   2895
      _Version        =   524298
      _ExtentX        =   5106
      _ExtentY        =   2408
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
      BackColorContainer=   2134465
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":D7FC
      textLT          =   "frmCharge.frx":D814
      textCT          =   "frmCharge.frx":D82C
      textRT          =   "frmCharge.frx":D844
      textLM          =   "frmCharge.frx":D85C
      textRM          =   "frmCharge.frx":D874
      textLB          =   "frmCharge.frx":D88C
      textCB          =   "frmCharge.frx":D8A4
      textRB          =   "frmCharge.frx":D8BC
      colorBack       =   "frmCharge.frx":D8D4
      colorIntern     =   "frmCharge.frx":D8FE
      colorMO         =   "frmCharge.frx":D928
      colorFocus      =   "frmCharge.frx":D952
      colorDisabled   =   "frmCharge.frx":D97C
      colorPressed    =   "frmCharge.frx":D9A6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdAcc 
      Height          =   8400
      Left            =   0
      TabIndex        =   23
      Top             =   270
      Visible         =   0   'False
      Width           =   45
      _cx             =   79
      _cy             =   14817
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
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1185
      Index           =   4
      Left            =   6210
      TabIndex        =   45
      Top             =   1290
      Width           =   1425
      _Version        =   524298
      _ExtentX        =   2514
      _ExtentY        =   2090
      _StockProps     =   66
      Caption         =   "Travel Agent"
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":D9D0
      textLT          =   "frmCharge.frx":DA48
      textCT          =   "frmCharge.frx":DA60
      textRT          =   "frmCharge.frx":DA78
      textLM          =   "frmCharge.frx":DA90
      textRM          =   "frmCharge.frx":DAA8
      textLB          =   "frmCharge.frx":DAC0
      textCB          =   "frmCharge.frx":DAD8
      textRB          =   "frmCharge.frx":DAF0
      colorBack       =   "frmCharge.frx":DB08
      colorIntern     =   "frmCharge.frx":DB32
      colorMO         =   "frmCharge.frx":DB5C
      colorFocus      =   "frmCharge.frx":DB86
      colorDisabled   =   "frmCharge.frx":DBB0
      colorPressed    =   "frmCharge.frx":DBDA
      Style           =   2
      Orientation     =   3
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1185
      Index           =   5
      Left            =   7620
      TabIndex        =   46
      Top             =   1290
      Width           =   1425
      _Version        =   524298
      _ExtentX        =   2514
      _ExtentY        =   2090
      _StockProps     =   66
      Caption         =   "Member"
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCharge.frx":DC04
      textLT          =   "frmCharge.frx":DC70
      textCT          =   "frmCharge.frx":DC88
      textRT          =   "frmCharge.frx":DCA0
      textLM          =   "frmCharge.frx":DCB8
      textRM          =   "frmCharge.frx":DCD0
      textLB          =   "frmCharge.frx":DCE8
      textCB          =   "frmCharge.frx":DD00
      textRB          =   "frmCharge.frx":DD18
      colorBack       =   "frmCharge.frx":DD30
      colorIntern     =   "frmCharge.frx":DD5A
      colorMO         =   "frmCharge.frx":DD84
      colorFocus      =   "frmCharge.frx":DDAE
      colorDisabled   =   "frmCharge.frx":DDD8
      colorPressed    =   "frmCharge.frx":DE02
      Style           =   2
      Orientation     =   3
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin VB.PictureBox picHoldFocus 
      Height          =   585
      Left            =   390
      ScaleHeight     =   525
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   1620
      Width           =   795
   End
   Begin MSForms.Label lblHeading 
      Height          =   525
      Left            =   480
      TabIndex        =   3
      Top             =   540
      Width           =   6675
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Select a Debtor or Room to Charge"
      Size            =   "11774;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   405
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click(Index As Integer)
    frmCharge.Tag = "Cancel"
    Me.Hide
End Sub
Private Sub cmdErr_Click()
    If picTender.Visible = True Then Exit Sub
    cmdErr.Caption = ""
    cmdErr.Visible = False
    errTimer.Enabled = False
    picHoldFocus.SetFocus
End Sub
Private Sub cmdInput_Click(Index As Integer)
    If picTender.Visible = True Then Exit Sub
    Select Case cmdInput(Index).Caption
        Case "Debtor"
            LoadDebtor
        Case "Staff"
            LoadStaff
        Case "Management"
            LoadManagement
        Case "Room"
            LoadRoom
        Case "Travel Agent"
            LoadTravel
        Case "Member"
            LoadMember
    End Select
    picHoldFocus.SetFocus
End Sub
Private Sub LoadTravel()
    grdAcc.Rows = 0
    cmdServers(0).Caption = ""
    cmdServers(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Debtors where Debt_Type = 3 order by Debtor_Name"
    If frmCharge.lblHeading.Tag = "Please Select a Room or Debtor" Then
        lblHeading.Caption = "Select a Travel Agent to Receive Payment on."
        lblHeading.Font.Size = 14
    Else
        lblHeading.Caption = "Select a Travel Agent to Charge to."
        lblHeading.Font.Size = 20
    End If
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdAcc.Rows = grdAcc.Rows + 1
        If i < 14 And Not rs.EOF Then
            cmdServers(i).Caption = rs.Fields("Debtor_No") & " - " & rs.Fields("Debtor_Name")
            cmdServers(i).Tag = rs.Fields("Debtor_No")
            If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
            grdAcc.Row = grdAcc.Rows - 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
        Else
            If b = 0 Then
                grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = "ßArrow"
                grdAcc.Rows = grdAcc.Rows + 1
                If i = 14 Then
                    cmdServers(14).Caption = ""
                    cmdServers(14).Tag = ""
                    cmdServers(14).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdServers(14).Visible = False Then cmdServers(14).Visible = True
                End If
            End If
            b = b + 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
            If b = 13 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdServers.Count - 1
       cmdServers(b).Caption = "0"
       cmdServers(b).Tag = "0"
       cmdServers(b).Visible = False
    Next b
End Sub
Private Sub LoadMember()
    grdAcc.Rows = 0
    cmdServers(0).Caption = ""
    cmdServers(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Debtors where Debt_Type = 4 order by Debtor_Name"
    If frmCharge.lblHeading.Tag = "Please Select a Room or Debtor" Then
        lblHeading.Caption = "Select a Member to Receive Payment on."
        lblHeading.Font.Size = 14
    Else
        lblHeading.Caption = "Select a Member to Charge to."
        lblHeading.Font.Size = 20
    End If
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdAcc.Rows = grdAcc.Rows + 1
        If i < 14 And Not rs.EOF Then
            cmdServers(i).Caption = rs.Fields("Debtor_No") & " - " & rs.Fields("Debtor_Name")
            cmdServers(i).Tag = rs.Fields("Debtor_No")
            If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
            grdAcc.Row = grdAcc.Rows - 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
        Else
            If b = 0 Then
                grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = "ßArrow"
                grdAcc.Rows = grdAcc.Rows + 1
                If i = 14 Then
                    cmdServers(14).Caption = ""
                    cmdServers(14).Tag = ""
                    cmdServers(14).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdServers(14).Visible = False Then cmdServers(14).Visible = True
                End If
            End If
            b = b + 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
            If b = 13 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdServers.Count - 1
       cmdServers(b).Caption = "0"
       cmdServers(b).Tag = "0"
       cmdServers(b).Visible = False
    Next b
End Sub
Private Sub LoadDebtor()
    grdAcc.Rows = 0
    cmdServers(0).Caption = ""
    cmdServers(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Debtors where Debt_Type = 0 order by Debtor_Name"
    If frmCharge.lblHeading.Tag = "Please Select a Room or Debtor" Then
        lblHeading.Caption = "Select a Debtor to Receive Payment on."
        lblHeading.Font.Size = 14
    Else
        lblHeading.Caption = "Select a Debtor to Charge to."
        lblHeading.Font.Size = 20
    End If
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdAcc.Rows = grdAcc.Rows + 1
        If i < 14 And Not rs.EOF Then
            cmdServers(i).Caption = rs.Fields("Debtor_No") & " - " & rs.Fields("Debtor_Name")
            cmdServers(i).Tag = rs.Fields("Debtor_No")
            If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
            grdAcc.Row = grdAcc.Rows - 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
        Else
            If b = 0 Then
                grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = "ßArrow"
                grdAcc.Rows = grdAcc.Rows + 1
                If i = 14 Then
                    cmdServers(14).Caption = ""
                    cmdServers(14).Tag = ""
                    cmdServers(14).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdServers(14).Visible = False Then cmdServers(14).Visible = True
                End If
            End If
            b = b + 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
            If b = 13 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdServers.Count - 1
       cmdServers(b).Caption = "0"
       cmdServers(b).Tag = "0"
       cmdServers(b).Visible = False
    Next b
End Sub
Private Sub LoadStaff()
    grdAcc.Rows = 0
    cmdServers(0).Caption = ""
    cmdServers(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Debtors where Debt_Type = 1  order by Debtor_Name"
    If frmCharge.lblHeading.Tag = "Please Select a Room or Debtor" Then
        lblHeading.Caption = "Select a Staff Account to Receive Payment on."
        lblHeading.Font.Size = 14
    Else
       lblHeading.Caption = "Select a Staff Account to Charge to."
       lblHeading.Font.Size = 20
    End If
    i = -1
    b = 0
While Not rs.EOF
        i = i + 1
        grdAcc.Rows = grdAcc.Rows + 1
        If i < 14 And Not rs.EOF Then
            cmdServers(i).Caption = rs.Fields("Debtor_No") & " - " & rs.Fields("Debtor_Name")
            cmdServers(i).Tag = rs.Fields("Debtor_No")
            If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
            grdAcc.Row = grdAcc.Rows - 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
        Else
            If b = 0 Then
                grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = "ßArrow"
                grdAcc.Rows = grdAcc.Rows + 1
                If i = 14 Then
                    cmdServers(14).Caption = ""
                    cmdServers(14).Tag = ""
                    cmdServers(14).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdServers(14).Visible = False Then cmdServers(14).Visible = True
                End If
            End If
            b = b + 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
            If b = 13 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdServers.Count - 1
       cmdServers(b).Caption = "0"
       cmdServers(b).Tag = "0"
       cmdServers(b).Visible = False
    Next b
End Sub
Private Sub LoadManagement()
    grdAcc.Rows = 0
    cmdServers(0).Caption = ""
    cmdServers(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Debtors where Debt_Type = 2 order by Debtor_Name"
    If frmCharge.lblHeading.Tag = "Please Select a Room or Debtor" Then
        lblHeading.Caption = "Select a Management Account to Receive Payment on."
        lblHeading.Font.Size = 14
    Else
        lblHeading.Caption = "Select a Management Account to Charge to."
        lblHeading.Font.Size = 20
    End If
    i = -1
    b = 0
   
   While Not rs.EOF
        i = i + 1
        grdAcc.Rows = grdAcc.Rows + 1
        If i < 14 And Not rs.EOF Then
            cmdServers(i).Caption = rs.Fields("Debtor_No") & " - " & rs.Fields("Debtor_Name")
            cmdServers(i).Tag = rs.Fields("Debtor_No")
            If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
            grdAcc.Row = grdAcc.Rows - 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
        Else
            If b = 0 Then
                grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = "ßArrow"
                grdAcc.Rows = grdAcc.Rows + 1
                If i = 14 Then
                    cmdServers(14).Caption = ""
                    cmdServers(14).Tag = ""
                    cmdServers(14).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdServers(14).Visible = False Then cmdServers(14).Visible = True
                End If
            End If
            b = b + 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Debtor_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Debtor_Name")
            If b = 13 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    For b = i + 1 To cmdServers.Count - 1
       cmdServers(b).Caption = "0"
       cmdServers(b).Tag = "0"
       cmdServers(b).Visible = False
    Next b
End Sub
Private Sub LoadRoom()
    grdAcc.Rows = 0
    cmdServers(0).Caption = ""
    cmdServers(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Rooms where Room_No in (Select Room_No from Reservations where Res_Type = 2)"
    If frmCharge.lblHeading.Tag = "Please Select a Room or Debtor" Then
        lblHeading.Caption = "Select a Guest to Receive Payment on."
        lblHeading.Font.Size = 14
    Else
        If picReceive.Visible = False Then lblHeading.Caption = "Select a Room to Charge to."
        lblHeading.Font.Size = 20
    End If
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdAcc.Rows = grdAcc.Rows + 1
        If i < 14 And Not rs.EOF Then
            cmdServers(i).Caption = rs.Fields("Room_No") & " - " & rs.Fields("Description")
            cmdServers(i).Tag = rs.Fields("Room_No")
            If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
            grdAcc.Row = grdAcc.Rows - 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Room_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Description")
        Else
            If b = 0 Then
                grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = "ßArrow"
                grdAcc.Rows = grdAcc.Rows + 1
                If i = 14 Then
                    cmdServers(14).Caption = ""
                    cmdServers(14).Tag = ""
                    cmdServers(14).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdServers(14).Visible = False Then cmdServers(14).Visible = True
                End If
            End If
            b = b + 1
            grdAcc.TextMatrix(grdAcc.Rows - 1, 0) = rs.Fields("Room_No")
            grdAcc.TextMatrix(grdAcc.Rows - 1, 1) = rs.Fields("Description")
            If b = 13 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdServers.Count - 1
       cmdServers(b).Caption = "0"
       cmdServers(b).Tag = "0"
       cmdServers(b).Visible = False
    Next b
End Sub

Private Sub cmdKey_Click(Index As Integer)
    picHoldFocus.SetFocus
    If picTender.Visible = True Then Exit Sub
    If errTimer.Enabled = True Then
        cmdErr.Caption = ""
        cmdErr.Visible = False
        errTimer.Enabled = False
    End If
    If cmdKey(Index).Caption = "CL" Then
        lblHeading.Caption = ""
        cmdErr.Caption = ""
        cmdErr.Visible = False
        errTimer.Enabled = False
        Exit Sub
    End If
    If cmdKey(Index).Caption = "OK" Then
        If Val(lblHeading.Caption) = 0 Then
            cmdErr.Caption = "Please Enter to Amount you are Receiving"
            cmdErr.Visible = True
            errTimer.Enabled = True
            lblHeading.Caption = ""
            Exit Sub
        End If
        If frmCharge.Tag <> "" Then
            frmCharge.Tag = frmCharge.Tag & ">" & lblHeading.Caption
        Else
            frmCharge.Tag = lblHeading.Caption & "> Deposit"
        End If
        cmdErr.Caption = "Please Select a Tender Type"
        cmdErr.Visible = True
        errTimer.Enabled = True
        picTender.Visible = True
        Exit Sub
    End If
    If lblHeading.Caption = "Please Enter the Amount Received on this Account" Then
         lblHeading.Caption = ""
         lblHeading.Font.Size = 20
    End If
    If InStr(lblHeading.Caption, "Deposit") <> 0 Then
        lblHeading.Caption = ""
    End If
    lblHeading.Caption = lblHeading.Caption & cmdKey(Index).Caption
End Sub

Private Sub cmdPrint_Click(Index As Integer)
    If picTender.Visible = True Then Exit Sub
End Sub

Private Sub cmdServers_Click(Index As Integer)
    DoEvents
    If cmdServers(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        grdAcc.Row = grdAcc.Row + 1
        For i = 0 To 14
            If grdAcc.TextMatrix(grdAcc.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdServers(i).Caption = ""
                    cmdServers(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
                Else
                    cmdServers(i).Caption = ""
                    cmdServers(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
                    grdAcc.Row = grdAcc.Row - 1
                    Exit For
                End If
            Else
                cmdServers(i).Caption = grdAcc.TextMatrix(grdAcc.Row, 0) & " - " & grdAcc.TextMatrix(grdAcc.Row, 1)
                cmdServers(i).Tag = grdAcc.TextMatrix(grdAcc.Row, 0)
            End If
            If grdAcc.Row = grdAcc.Rows - 1 Then Exit For
            grdAcc.Row = grdAcc.Row + 1
        Next i
        For b = i + 1 To cmdServers.Count - 1
            cmdServers(b).Caption = "1"
            cmdServers(b).Tag = ""
            cmdServers(b).Visible = False
        Next b
        Exit Sub
        
    End If
    If cmdServers(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdServers(0).Picture = ""
        While grdAcc.TextMatrix(grdAcc.Row, 0) <> "ßArrow"
            grdAcc.Row = grdAcc.Row - 1
        Wend
        grdAcc.Row = grdAcc.Row - 14
        For i = 0 To 14
            If grdAcc.TextMatrix(grdAcc.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdServers(i).Caption = ""
                    cmdServers(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
                Else
                    cmdServers(i).Caption = ""
                    cmdServers(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
                    grdAcc.Row = grdAcc.Row - 1
                    Exit For
                End If
            Else
                cmdServers(i).Caption = grdAcc.TextMatrix(grdAcc.Row, 0) & " - " & grdAcc.TextMatrix(grdAcc.Row, 1)
                cmdServers(i).Tag = grdAcc.TextMatrix(grdAcc.Row, 0)
                If cmdServers(i).Visible = False Then cmdServers(i).Visible = True
            End If
            If grdAcc.Row = grdAcc.Rows - 1 Then Exit For
            grdAcc.Row = grdAcc.Row + 1
        Next i
        For b = i + 1 To cmdServers.Count - 1
            cmdServers(b).Caption = "1"
            cmdServers(b).Tag = ""
            cmdServers(b).Visible = False
        Next b
        Exit Sub
    End If
    If frmCharge.lblHeading.Tag = "Please Select a Room or Debtor" Then
        frmCharge.Tag = cmdServers(Index).Tag & " > " & lblHeading.Caption
        cmdInput(0).Enabled = False
        cmdInput(1).Enabled = False
        cmdInput(2).Enabled = False
        cmdInput(3).Enabled = False
        If InStr(frmCharge.Tag, "Guest") <> 0 Then
            ActiveReadServer "Select Res_No from Reservations where Res_Type = 2 and Room_No in (Select Room_No from Reservations where Room_No = " & cmdServers(Index).Tag & " and Res_Type=2)"
            TillData.Res_No = rs.Fields("Res_No")
            rs.Close
            DoEvents
            ActiveReadServer "Select * from Room_Accounts where Res_No = " & TillData.Res_No & " order by Line_No"
        Else
            ActiveReadServer "Select * from Debtor_Accounts where Account_No = '" & cmdServers(Index).Tag & "' order by Line_No"
        End If
        picReceive.Visible = True
        grdMain.Rows = 1
        Balance = 0
        While Not rs.EOF
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            If rs.Fields("Transaction_Type") & "" = "Accomodation" Then
                grdMain.TextMatrix(grdMain.Row, 0) = "Accomodation"
                grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
                Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
            End If
            If rs.Fields("Transaction_Type") & "" = "Invoice" Then
                grdMain.TextMatrix(grdMain.Row, 0) = "Sales Invoice"
                grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
                Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
            End If
            If rs.Fields("Transaction_Type") & "" = "Receipt" Then
                grdMain.TextMatrix(grdMain.Row, 0) = "Receive on Account"
                grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Credit") * -1 & "", "0.00")
                Balance = Balance + Format(rs.Fields("Credit") * -1 & "", "0.00")
            End If
            If rs.Fields("Transaction_Type") & "" = "Journal" Then
                grdMain.TextMatrix(grdMain.Row, 0) = "Journal" & " - " & rs.Fields("Ref_No")
                If rs.Fields("Credit") <> 0 Then
                    grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Credit") * -1 & "", "0.00")
                    Balance = Balance + Format(rs.Fields("Credit") * -1 & "", "0.00")
                End If
                If rs.Fields("Debit") <> 0 Then
                    grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
                    Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
                End If
            End If
            If rs.Fields("Transaction_Type") & "" = "Deposit" Then
                grdMain.TextMatrix(grdMain.Row, 0) = "Deposit Received"
                grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Credit") * -1 & "", "0.00")
                Balance = Balance + Format(rs.Fields("Credit") * -1 & "", "0.00")
            End If
            If Left(rs.Fields("Transaction_Type") & "", 9) = "Telephone" Then
                grdMain.TextMatrix(grdMain.Row, 0) = "Telephone Charge"
                grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
                Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
            End If
            rs.MoveNext
        Wend
        grdMain.ShowCell grdMain.Row, 0
        rs.Close
        lblTender.Caption = Format(Balance, "0.00")
        frmCharge.lblHeading.Caption = "Please Enter the Amount Received on this Account"
    Else
        If frmCharge.lblHeading.Tag = "Select a Debtor to Close the Room to." Then
            ActiveReadServer "Select * from Room_Accounts where Res_No = " & TillData.Res_No & " order by Line_No"
            While Not rs.EOF
                Balance = 0
                ActiveReadServer1 "Select Balance from Debtors where Debtor_No = '" & cmdServers(Index).Tag & "'"
                If rs1.RecordCount > 0 Then
                    Balance = Balance + (rs.Fields("Debit") - rs.Fields("Credit"))
                End If
                rs1.Close
                ActiveUpdateServer "Insert into Debtor_Accounts (Date_Time,Invoice_No,Transaction_Type,Account_No,Debit,Credit,Balance) values ('" & rs.Fields("Date_Time") & "','" & rs.Fields("Invoice_No") & "','" & rs.Fields("Transaction_Type") & "','" & cmdServers(Index).Tag & "','" & rs.Fields("Debit") & "','" & rs.Fields("Credit") & "'," & Balance & ")"
                DoEvents
                ActiveUpdateServer "Update Debtors set Balance = " & Balance & " where Debtor_No = '" & cmdServers(Index).Tag & "'"
                DoEvents
                rs.MoveNext
            Wend
            rs.Close
            ActiveReadServer1 "Select isnull(max(Invoice_No),0)+1 as Receipt_No from Room_Accounts where Transaction_Type = 'Receipt'"
            Deposit_No = rs1.Fields("Receipt_No")
            rs1.Close
            Balance = 0
            ActiveReadServer "Select * from Room_Accounts where Res_No = '" & TillData.Res_No & "' order by Line_No"
            While Not rs.EOF
                Balance = Balance + (rs.Fields("Debit") - rs.Fields("Credit"))
                rs.MoveNext
            Wend
            rs.Close
            TillData.Change = 0
            Deposit = Balance
            
            ActiveUpdateServer "INSERT INTO [Room_Accounts]([User_No],[Date_Time],[Transaction_Type], [Invoice_No], [Account_No], [Res_No], [Debit], [Credit], [Balance],[Tender_Type])" & _
            "VALUES(" & UserRecord.User_Number & ",Getdate(),'Receipt'," & Deposit_No & ",'" & TillData.Room_No & "','" & TillData.Res_No & "',0," & Deposit & "," & Balance + (Deposit * -1) & ",'Charge')"
            DoEvents
            
            ActiveUpdateServer "Update Reservations set Res_Type = 3 where Res_No=" & TillData.Res_No
            Unload Me
            Exit Sub
        Else
            ActiveReadServer "Select Credit_Limit,Balance from Debtors where Debtor_No ='" & cmdServers(Index).Tag & "'"
            If rs.RecordCount > 0 Then
                If Val(rs.Fields("Credit_Limit") & "") <> 0 Then
                    If TillData.SaleTotal + Val(rs.Fields("Balance") & "") > Val(rs.Fields("Credit_Limit") & "") Then
                        cmdErr.Caption = "Credit Limit Exceeded for this Account"
                        cmdErr.Visible = True
                        errTimer.Enabled = True
                        lblHeading.Caption = ""
                        rs.Close
                        Exit Sub
                    End If
                End If
            End If
            rs.Close
            frmCharge.Tag = cmdServers(Index).Tag & " > " & lblHeading.Caption
        End If
        Me.Hide
    End If
End Sub
Private Sub cmdTender_Click(Index As Integer)
    frmCharge.Tag = frmCharge.Tag & "|" & cmdTender(Index).Caption
    Me.Hide
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
    frmCharge.KeyPreview = True
    picHoldFocus.SetFocus
    Screen.MousePointer = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static LoadDebtor As String
    If Shift <> 0 Then
        KeyCode = 0
        Exit Sub
    End If
    buttonNo = -1
    If KeyCode = 13 Then
        KeyCode = 0
        For i = 0 To cmdInput.Count - 1
            If cmdInput(i).Value <> 0 Then
                buttonNo = i
                Exit For
            End If
        Next i
        If buttonNo = -1 Then
            cmdErr.Caption = "Select a Debtor Type before Scanning"
            cmdErr.Visible = True
            errTimer.Enabled = True
            lblHeading.Caption = ""
            LoadDebtor = ""
            Exit Sub
        End If
        Select Case cmdInput(buttonNo).Caption
            Case "Debtor"
                DebtType = 0
            Case "Staff"
                DebtType = 1
            Case "Management"
                DebtType = 2
            Case "Room"
            Case "Travel Agent"
                DebtType = 3
            Case "Member"
                DebtType = 4
        End Select
        ActiveReadServer "Select * from Debtors where Debt_Type = " & DebtType & " and Debtor_No = '" & LoadDebtor & "'"
        If rs.RecordCount > 0 Then
            rs.Close
            If frmCharge.lblHeading.Tag = "Please Select a Room or Debtor" Then
                frmCharge.Tag = LoadDebtor & " > " & lblHeading.Caption
                cmdInput(0).Enabled = False
                cmdInput(1).Enabled = False
                cmdInput(2).Enabled = False
                cmdInput(3).Enabled = False
                If InStr(frmCharge.Tag, "Guest") <> 0 Then
                    ActiveReadServer "Select Res_No from Reservations where Res_Type = 2 and Room_No in (Select Room_No from Reservations where Room_No = " & LoadDebtor & " and Res_Type=2)"
                    TillData.Res_No = rs.Fields("Res_No")
                    rs.Close
                    DoEvents
                    ActiveReadServer "Select * from Room_Accounts where Res_No = " & TillData.Res_No & " order by Line_No"
                Else
                    ActiveReadServer "Select * from Debtor_Accounts where Account_No = '" & LoadDebtor & "' order by Line_No"
                End If
                picReceive.Visible = True
                grdMain.Rows = 1
                Balance = 0
                While Not rs.EOF
                    grdMain.Rows = grdMain.Rows + 1
                    grdMain.Row = grdMain.Rows - 1
                    If rs.Fields("Transaction_Type") & "" = "Accomodation" Then
                        grdMain.TextMatrix(grdMain.Row, 0) = "Accomodation"
                        grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
                        Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
                    End If
                    If rs.Fields("Transaction_Type") & "" = "Invoice" Then
                        grdMain.TextMatrix(grdMain.Row, 0) = "Sales Invoice"
                        grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
                        Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
                    End If
                    If rs.Fields("Transaction_Type") & "" = "Receipt" Then
                        grdMain.TextMatrix(grdMain.Row, 0) = "Receive on Account"
                        grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Credit") * -1 & "", "0.00")
                        Balance = Balance + Format(rs.Fields("Credit") * -1 & "", "0.00")
                    End If
                    If rs.Fields("Transaction_Type") & "" = "Journal" Then
                        grdMain.TextMatrix(grdMain.Row, 0) = "Journal" & " - " & rs.Fields("Ref_No")
                        If rs.Fields("Credit") <> 0 Then
                            grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Credit") * -1 & "", "0.00")
                            Balance = Balance + Format(rs.Fields("Credit") * -1 & "", "0.00")
                        End If
                        If rs.Fields("Debit") <> 0 Then
                            grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
                            Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
                        End If
                    End If
                    If rs.Fields("Transaction_Type") & "" = "Deposit" Then
                        grdMain.TextMatrix(grdMain.Row, 0) = "Deposit Received"
                        grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Credit") * -1 & "", "0.00")
                        Balance = Balance + Format(rs.Fields("Credit") * -1 & "", "0.00")
                    End If
                    If Left(rs.Fields("Transaction_Type") & "", 9) = "Telephone" Then
                        grdMain.TextMatrix(grdMain.Row, 0) = "Telephone Charge"
                        grdMain.TextMatrix(grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
                        Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
                    End If
                    rs.MoveNext
                Wend
                grdMain.ShowCell grdMain.Row, 0
                rs.Close
                lblTender.Caption = Format(Balance, "0.00")
                frmCharge.lblHeading.Caption = "Please Enter the Amount Received on this Account"
            Else
                If frmCharge.lblHeading.Tag = "Select a Debtor to Close the Room to." Then
                    ActiveReadServer "Select * from Room_Accounts where Res_No = " & TillData.Res_No & " order by Line_No"
                    While Not rs.EOF
                        Balance = 0
                        ActiveReadServer1 "Select Balance from Debtors where Debtor_No = '" & LoadDebtor & "'"
                        If rs1.RecordCount > 0 Then
                            Balance = Balance + (rs.Fields("Debit") - rs.Fields("Credit"))
                        End If
                        rs1.Close
                        ActiveUpdateServer "Insert into Debtor_Accounts (Date_Time,Invoice_No,Transaction_Type,Account_No,Debit,Credit,Balance) values ('" & rs.Fields("Date_Time") & "','" & rs.Fields("Invoice_No") & "','" & rs.Fields("Transaction_Type") & "','" & LoadDebtor & "','" & rs.Fields("Debit") & "','" & rs.Fields("Credit") & "'," & Balance & ")"
                        DoEvents
                        ActiveUpdateServer "Update Debtors set Balance = " & Balance & " where Debtor_No = '" & LoadDebtor & "'"
                        DoEvents
                        rs.MoveNext
                    Wend
                    rs.Close
                    ActiveReadServer1 "Select isnull(max(Invoice_No),0)+1 as Receipt_No from Room_Accounts where Transaction_Type = 'Receipt'"
                    Deposit_No = rs1.Fields("Receipt_No")
                    rs1.Close
                    Balance = 0
                    ActiveReadServer "Select * from Room_Accounts where Res_No = '" & TillData.Res_No & "' order by Line_No"
                    While Not rs.EOF
                        Balance = Balance + (rs.Fields("Debit") - rs.Fields("Credit"))
                        rs.MoveNext
                    Wend
                    rs.Close
                    TillData.Change = 0
                    Deposit = Balance
                    
                    ActiveUpdateServer "INSERT INTO [Room_Accounts]([User_No],[Date_Time],[Transaction_Type], [Invoice_No], [Account_No], [Res_No], [Debit], [Credit], [Balance],[Tender_Type])" & _
                    "VALUES(" & UserRecord.User_Number & ",Getdate(),'Receipt'," & Deposit_No & ",'" & TillData.Room_No & "','" & TillData.Res_No & "',0," & Deposit & "," & Balance + (Deposit * -1) & ",'Charge')"
                    DoEvents
                    
                    ActiveUpdateServer "Update Reservations set Res_Type = 3 where Res_No=" & TillData.Res_No
                    Unload Me
                    Exit Sub
                Else
                    ActiveReadServer "Select Credit_Limit,Balance from Debtors where Debtor_No ='" & LoadDebtor & "'"
                    If rs.RecordCount > 0 Then
                        If Val(rs.Fields("Credit_Limit") & "") <> 0 Then
                            If TillData.SaleTotal + Val(rs.Fields("Balance") & "") > Val(rs.Fields("Credit_Limit") & "") Then
                                cmdErr.Caption = "Credit Limit Exceeded for this Account"
                                cmdErr.Visible = True
                                errTimer.Enabled = True
                                lblHeading.Caption = ""
                                rs.Close
                                Exit Sub
                            End If
                        End If
                    End If
                    rs.Close
                    frmCharge.Tag = LoadDebtor & " > " & lblHeading.Caption
                End If
                Me.Hide
            End If
            LoadDebtor = ""
            Exit Sub
        Else
            cmdErr.Caption = "Unknown " & cmdInput(buttonNo).Caption
            cmdErr.Visible = True
            errTimer.Enabled = True
            lblHeading.Caption = ""
        End If
        rs.Close
        LoadDebtor = ""
    Else
        If KeyCode = 8 Then
            LoadDebtor = ""
            KeyCode = 0
        End If
        If KeyCode > 47 And KeyCode < 58 Then
            LoadDebtor = LoadDebtor & Chr(KeyCode)
            KeyCode = 0
        End If
        If KeyCode > 95 And KeyCode < 106 Then
            Select Case KeyCode
                Case 96: LoadDebtor = LoadDebtor & "0"
                Case 97: LoadDebtor = LoadDebtor & "1"
                Case 98: LoadDebtor = LoadDebtor & "2"
                Case 99: LoadDebtor = LoadDebtor & "3"
                Case 100: LoadDebtor = LoadDebtor & "4"
                Case 101: LoadDebtor = LoadDebtor & "5"
                Case 101: LoadDebtor = LoadDebtor & "6"
                Case 103: LoadDebtor = LoadDebtor & "7"
                Case 104: LoadDebtor = LoadDebtor & "8"
                Case 104: LoadDebtor = LoadDebtor & "9"
            End Select
            KeyCode = 0
        End If
    End If
End Sub
Private Sub Form_Load()
    picTender.Visible = False
    grdMain.TextMatrix(0, 0) = "Transaction"
    grdMain.TextMatrix(0, 1) = "Value"
    grdMain.ColWidth(0) = grdMain.Width * 0.7
    grdMain.ColWidth(1) = grdMain.Width * 0.3
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignRightCenter
    lblHeading.Font.Size = 20
End Sub

