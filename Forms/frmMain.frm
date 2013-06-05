VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10650
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14085
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picAccBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   475
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   14085
      TabIndex        =   25
      Top             =   1095
      Visible         =   0   'False
      Width           =   14085
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   9
         Left            =   2820
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Width           =   2805
         _Version        =   524298
         _ExtentX        =   4939
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Staff Accounts"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":1CCA
         textLT          =   "frmMain.frx":1D46
         textCT          =   "frmMain.frx":1D5E
         textRT          =   "frmMain.frx":1D76
         textLM          =   "frmMain.frx":1D8E
         textRM          =   "frmMain.frx":1DA6
         textLB          =   "frmMain.frx":1DBE
         textCB          =   "frmMain.frx":1DD6
         textRB          =   "frmMain.frx":1DEE
         colorBack       =   "frmMain.frx":1E06
         colorIntern     =   "frmMain.frx":1E30
         colorMO         =   "frmMain.frx":1E5A
         colorFocus      =   "frmMain.frx":1E84
         colorDisabled   =   "frmMain.frx":1EAE
         colorPressed    =   "frmMain.frx":1ED8
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   10
         Left            =   5640
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   0
         Width           =   2805
         _Version        =   524298
         _ExtentX        =   4939
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Management Accounts"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":1F02
         textLT          =   "frmMain.frx":1F88
         textCT          =   "frmMain.frx":1FA0
         textRT          =   "frmMain.frx":1FB8
         textLM          =   "frmMain.frx":1FD0
         textRM          =   "frmMain.frx":1FE8
         textLB          =   "frmMain.frx":2000
         textCB          =   "frmMain.frx":2018
         textRB          =   "frmMain.frx":2030
         colorBack       =   "frmMain.frx":2048
         colorIntern     =   "frmMain.frx":2072
         colorMO         =   "frmMain.frx":209C
         colorFocus      =   "frmMain.frx":20C6
         colorDisabled   =   "frmMain.frx":20F0
         colorPressed    =   "frmMain.frx":211A
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   12
         Left            =   8460
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   0
         Width           =   2805
         _Version        =   524298
         _ExtentX        =   4939
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Travel Agents"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":2144
         textLT          =   "frmMain.frx":21BE
         textCT          =   "frmMain.frx":21D6
         textRT          =   "frmMain.frx":21EE
         textLM          =   "frmMain.frx":2206
         textRM          =   "frmMain.frx":221E
         textLB          =   "frmMain.frx":2236
         textCB          =   "frmMain.frx":224E
         textRB          =   "frmMain.frx":2266
         colorBack       =   "frmMain.frx":227E
         colorIntern     =   "frmMain.frx":22A8
         colorMO         =   "frmMain.frx":22D2
         colorFocus      =   "frmMain.frx":22FC
         colorDisabled   =   "frmMain.frx":2326
         colorPressed    =   "frmMain.frx":2350
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   14
         Left            =   11280
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Width           =   2715
         _Version        =   524298
         _ExtentX        =   4789
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Members"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":237A
         textLT          =   "frmMain.frx":23E8
         textCT          =   "frmMain.frx":2400
         textRT          =   "frmMain.frx":2418
         textLM          =   "frmMain.frx":2430
         textRM          =   "frmMain.frx":2448
         textLB          =   "frmMain.frx":2460
         textCB          =   "frmMain.frx":2478
         textRB          =   "frmMain.frx":2490
         colorBack       =   "frmMain.frx":24A8
         colorIntern     =   "frmMain.frx":24D2
         colorMO         =   "frmMain.frx":24FC
         colorFocus      =   "frmMain.frx":2526
         colorDisabled   =   "frmMain.frx":2550
         colorPressed    =   "frmMain.frx":257A
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   15
         Left            =   14010
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   2715
         _Version        =   524298
         _ExtentX        =   4789
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Email"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":25A4
         textLT          =   "frmMain.frx":260E
         textCT          =   "frmMain.frx":2626
         textRT          =   "frmMain.frx":263E
         textLM          =   "frmMain.frx":2656
         textRM          =   "frmMain.frx":266E
         textLB          =   "frmMain.frx":2686
         textCB          =   "frmMain.frx":269E
         textRB          =   "frmMain.frx":26B6
         colorBack       =   "frmMain.frx":26CE
         colorIntern     =   "frmMain.frx":26F8
         colorMO         =   "frmMain.frx":2722
         colorFocus      =   "frmMain.frx":274C
         colorDisabled   =   "frmMain.frx":2776
         colorPressed    =   "frmMain.frx":27A0
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   11
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   0
         Width           =   2800
         _Version        =   524298
         _ExtentX        =   4939
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Debtors"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":27CA
         textLT          =   "frmMain.frx":2838
         textCT          =   "frmMain.frx":2850
         textRT          =   "frmMain.frx":2868
         textLM          =   "frmMain.frx":2880
         textRM          =   "frmMain.frx":2898
         textLB          =   "frmMain.frx":28B0
         textCB          =   "frmMain.frx":28C8
         textRB          =   "frmMain.frx":28E0
         colorBack       =   "frmMain.frx":28F8
         colorIntern     =   "frmMain.frx":2922
         colorMO         =   "frmMain.frx":294C
         colorFocus      =   "frmMain.frx":2976
         colorDisabled   =   "frmMain.frx":29A0
         colorPressed    =   "frmMain.frx":29CA
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
         Value           =   -1  'True
      End
   End
   Begin VB.PictureBox picProdBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   14085
      TabIndex        =   13
      Top             =   1575
      Visible         =   0   'False
      Width           =   14085
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   1
         Left            =   1605
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1650
         _Version        =   524298
         _ExtentX        =   2910
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Departments"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         CaptionWordWrapPerc=   94
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":29F4
         textLT          =   "frmMain.frx":2A6A
         textCT          =   "frmMain.frx":2A82
         textRT          =   "frmMain.frx":2A9A
         textLM          =   "frmMain.frx":2AB2
         textRM          =   "frmMain.frx":2ACA
         textLB          =   "frmMain.frx":2AE2
         textCB          =   "frmMain.frx":2AFA
         textRB          =   "frmMain.frx":2B12
         colorBack       =   "frmMain.frx":2B2A
         colorIntern     =   "frmMain.frx":2B54
         colorMO         =   "frmMain.frx":2B7E
         colorFocus      =   "frmMain.frx":2BA8
         colorDisabled   =   "frmMain.frx":2BD2
         colorPressed    =   "frmMain.frx":2BFC
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   3
         Left            =   4755
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   1470
         _Version        =   524298
         _ExtentX        =   2593
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Suppliers"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":2C26
         textLT          =   "frmMain.frx":2C98
         textCT          =   "frmMain.frx":2CB0
         textRT          =   "frmMain.frx":2CC8
         textLM          =   "frmMain.frx":2CE0
         textRM          =   "frmMain.frx":2CF8
         textLB          =   "frmMain.frx":2D10
         textCB          =   "frmMain.frx":2D28
         textRB          =   "frmMain.frx":2D40
         colorBack       =   "frmMain.frx":2D58
         colorIntern     =   "frmMain.frx":2D82
         colorMO         =   "frmMain.frx":2DAC
         colorFocus      =   "frmMain.frx":2DD6
         colorDisabled   =   "frmMain.frx":2E00
         colorPressed    =   "frmMain.frx":2E2A
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenuSep 
         Height          =   465
         Left            =   6240
         TabIndex        =   16
         Top             =   0
         Width           =   135
         _Version        =   524298
         _ExtentX        =   238
         _ExtentY        =   820
         _StockProps     =   66
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         Clickable       =   0   'False
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":2E54
         textLT          =   "frmMain.frx":2E6C
         textCT          =   "frmMain.frx":2E84
         textRT          =   "frmMain.frx":2E9C
         textLM          =   "frmMain.frx":2EB4
         textRM          =   "frmMain.frx":2ECC
         textLB          =   "frmMain.frx":2EE4
         textCB          =   "frmMain.frx":2EFC
         textRB          =   "frmMain.frx":2F14
         colorBack       =   "frmMain.frx":2F2C
         colorIntern     =   "frmMain.frx":2F56
         colorMO         =   "frmMain.frx":2F80
         colorFocus      =   "frmMain.frx":2FAA
         colorDisabled   =   "frmMain.frx":2FD4
         colorPressed    =   "frmMain.frx":2FFE
         Style           =   7
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   4
         Left            =   6390
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   1695
         _Version        =   524298
         _ExtentX        =   2990
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Orders"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":3028
         textLT          =   "frmMain.frx":3094
         textCT          =   "frmMain.frx":30AC
         textRT          =   "frmMain.frx":30C4
         textLM          =   "frmMain.frx":30DC
         textRM          =   "frmMain.frx":30F4
         textLB          =   "frmMain.frx":310C
         textCB          =   "frmMain.frx":3124
         textRB          =   "frmMain.frx":313C
         colorBack       =   "frmMain.frx":3154
         colorIntern     =   "frmMain.frx":317E
         colorMO         =   "frmMain.frx":31A8
         colorFocus      =   "frmMain.frx":31D2
         colorDisabled   =   "frmMain.frx":31FC
         colorPressed    =   "frmMain.frx":3226
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   5
         Left            =   8160
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   2340
         _Version        =   524298
         _ExtentX        =   4128
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Goods Receiving"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         CaptionWordWrapPerc=   95
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":3250
         textLT          =   "frmMain.frx":32CE
         textCT          =   "frmMain.frx":32E6
         textRT          =   "frmMain.frx":32FE
         textLM          =   "frmMain.frx":3316
         textRM          =   "frmMain.frx":332E
         textLB          =   "frmMain.frx":3346
         textCB          =   "frmMain.frx":335E
         textRB          =   "frmMain.frx":3376
         colorBack       =   "frmMain.frx":338E
         colorIntern     =   "frmMain.frx":33B8
         colorMO         =   "frmMain.frx":33E2
         colorFocus      =   "frmMain.frx":340C
         colorDisabled   =   "frmMain.frx":3436
         colorPressed    =   "frmMain.frx":3460
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   6
         Left            =   10560
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   2115
         _Version        =   524298
         _ExtentX        =   3731
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Transfers"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":348A
         textLT          =   "frmMain.frx":34FC
         textCT          =   "frmMain.frx":3514
         textRT          =   "frmMain.frx":352C
         textLM          =   "frmMain.frx":3544
         textRM          =   "frmMain.frx":355C
         textLB          =   "frmMain.frx":3574
         textCB          =   "frmMain.frx":358C
         textRB          =   "frmMain.frx":35A4
         colorBack       =   "frmMain.frx":35BC
         colorIntern     =   "frmMain.frx":35E6
         colorMO         =   "frmMain.frx":3610
         colorFocus      =   "frmMain.frx":363A
         colorDisabled   =   "frmMain.frx":3664
         colorPressed    =   "frmMain.frx":368E
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   7
         Left            =   14490
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         _Version        =   524298
         _ExtentX        =   2990
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Stock Takes"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":36B8
         textLT          =   "frmMain.frx":372E
         textCT          =   "frmMain.frx":3746
         textRT          =   "frmMain.frx":375E
         textLM          =   "frmMain.frx":3776
         textRM          =   "frmMain.frx":378E
         textLB          =   "frmMain.frx":37A6
         textCB          =   "frmMain.frx":37BE
         textRB          =   "frmMain.frx":37D6
         colorBack       =   "frmMain.frx":37EE
         colorIntern     =   "frmMain.frx":3818
         colorMO         =   "frmMain.frx":3842
         colorFocus      =   "frmMain.frx":386C
         colorDisabled   =   "frmMain.frx":3896
         colorPressed    =   "frmMain.frx":38C0
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   0
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   1710
         _Version        =   524298
         _ExtentX        =   3016
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Products"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":38EA
         textLT          =   "frmMain.frx":395A
         textCT          =   "frmMain.frx":3972
         textRT          =   "frmMain.frx":398A
         textLM          =   "frmMain.frx":39A2
         textRM          =   "frmMain.frx":39BA
         textLB          =   "frmMain.frx":39D2
         textCB          =   "frmMain.frx":39EA
         textRB          =   "frmMain.frx":3A02
         colorBack       =   "frmMain.frx":3A1A
         colorIntern     =   "frmMain.frx":3A44
         colorMO         =   "frmMain.frx":3A6E
         colorFocus      =   "frmMain.frx":3A98
         colorDisabled   =   "frmMain.frx":3AC2
         colorPressed    =   "frmMain.frx":3AEC
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
         Value           =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   8
         Left            =   9735
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1530
         _Version        =   524298
         _ExtentX        =   2699
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "BackOrder"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":3B16
         textLT          =   "frmMain.frx":3B88
         textCT          =   "frmMain.frx":3BA0
         textRT          =   "frmMain.frx":3BB8
         textLM          =   "frmMain.frx":3BD0
         textRM          =   "frmMain.frx":3BE8
         textLB          =   "frmMain.frx":3C00
         textCB          =   "frmMain.frx":3C18
         textRB          =   "frmMain.frx":3C30
         colorBack       =   "frmMain.frx":3C48
         colorIntern     =   "frmMain.frx":3C72
         colorMO         =   "frmMain.frx":3C9C
         colorFocus      =   "frmMain.frx":3CC6
         colorDisabled   =   "frmMain.frx":3CF0
         colorPressed    =   "frmMain.frx":3D1A
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   2
         Left            =   3270
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   0
         Width           =   1470
         _Version        =   524298
         _ExtentX        =   2593
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Locations"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":3D44
         textLT          =   "frmMain.frx":3DB6
         textCT          =   "frmMain.frx":3DCE
         textRT          =   "frmMain.frx":3DE6
         textLM          =   "frmMain.frx":3DFE
         textRM          =   "frmMain.frx":3E16
         textLB          =   "frmMain.frx":3E2E
         textCB          =   "frmMain.frx":3E46
         textRB          =   "frmMain.frx":3E5E
         colorBack       =   "frmMain.frx":3E76
         colorIntern     =   "frmMain.frx":3EA0
         colorMO         =   "frmMain.frx":3ECA
         colorFocus      =   "frmMain.frx":3EF4
         colorDisabled   =   "frmMain.frx":3F1E
         colorPressed    =   "frmMain.frx":3F48
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
      Begin BTNENHLib4.BtnEnh cmdMenu 
         Height          =   465
         Index           =   13
         Left            =   12690
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   0
         Width           =   1785
         _Version        =   524298
         _ExtentX        =   3149
         _ExtentY        =   820
         _StockProps     =   66
         Caption         =   "Production"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Surface         =   10
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmMain.frx":3F72
         textLT          =   "frmMain.frx":3FE6
         textCT          =   "frmMain.frx":3FFE
         textRT          =   "frmMain.frx":4016
         textLM          =   "frmMain.frx":402E
         textRM          =   "frmMain.frx":4046
         textLB          =   "frmMain.frx":405E
         textCB          =   "frmMain.frx":4076
         textRB          =   "frmMain.frx":408E
         colorBack       =   "frmMain.frx":40A6
         colorIntern     =   "frmMain.frx":40D0
         colorMO         =   "frmMain.frx":40FA
         colorFocus      =   "frmMain.frx":4124
         colorDisabled   =   "frmMain.frx":414E
         colorPressed    =   "frmMain.frx":4178
         Style           =   2
         Orientation     =   2
         HollowFrame     =   -1  'True
         LightDirection  =   1
         RectHardEdges   =   -1  'True
      End
   End
   Begin VB.PictureBox picSide 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8235
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   1605
      TabIndex        =   3
      Top             =   2070
      Width           =   1605
      Begin btButtonEx.ButtonEx cmdBar 
         Height          =   1050
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Tag             =   "Reservations"
         Top             =   150
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Appearance      =   4
         AutoMask        =   0   'False
         BackColor       =   -2147483636
         Caption         =   ""
         CaptionOffsetY  =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   8421504
         SkinDisabled    =   "frmMain.frx":41A2
         SkinOver        =   "frmMain.frx":7F3C
         SkinUp          =   "frmMain.frx":BCD6
         MaskColor       =   8421504
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdBar 
         Height          =   1050
         Index           =   1
         Left            =   270
         TabIndex        =   5
         Tag             =   "Rooms"
         Top             =   2310
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Appearance      =   4
         AutoMask        =   0   'False
         BackColor       =   -2147483636
         Caption         =   ""
         CaptionOffsetY  =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   8421504
         SkinDisabled    =   "frmMain.frx":FA70
         SkinOver        =   "frmMain.frx":1380A
         SkinUp          =   "frmMain.frx":175A4
         MaskColor       =   8421504
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdBar 
         Height          =   1050
         Index           =   2
         Left            =   270
         TabIndex        =   6
         Tag             =   "Guests"
         Top             =   3390
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Appearance      =   4
         AutoMask        =   0   'False
         BackColor       =   -2147483636
         Caption         =   ""
         CaptionOffsetY  =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   8421504
         SkinDisabled    =   "frmMain.frx":1B33E
         SkinDown        =   "frmMain.frx":1F0D8
         SkinOver        =   "frmMain.frx":22E72
         SkinUp          =   "frmMain.frx":26C0C
         MaskColor       =   8421504
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdBar 
         Height          =   1050
         Index           =   3
         Left            =   270
         TabIndex        =   7
         Tag             =   "Checkin"
         Top             =   4470
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Appearance      =   4
         AutoMask        =   0   'False
         Enabled         =   0   'False
         BackColor       =   -2147483636
         Caption         =   ""
         CaptionOffsetY  =   10
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   8421504
         SkinDisabled    =   "frmMain.frx":2A9A6
         SkinOver        =   "frmMain.frx":2E740
         SkinUp          =   "frmMain.frx":324DA
         MaskColor       =   8421504
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdBar 
         Height          =   1050
         Index           =   4
         Left            =   270
         TabIndex        =   8
         Tag             =   "Users"
         Top             =   6630
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Appearance      =   4
         AutoMask        =   0   'False
         BackColor       =   -2147483636
         Caption         =   ""
         CaptionOffsetY  =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   8421504
         SkinDisabled    =   "frmMain.frx":36274
         SkinOver        =   "frmMain.frx":3A00E
         SkinUp          =   "frmMain.frx":3DDA8
         MaskColor       =   8421504
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdBar 
         Height          =   1050
         Index           =   5
         Left            =   300
         TabIndex        =   9
         Tag             =   "Reports"
         Top             =   7710
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Appearance      =   4
         AutoMask        =   0   'False
         BackColor       =   -2147483636
         Caption         =   ""
         CaptionOffsetY  =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   8421504
         SkinDisabled    =   "frmMain.frx":41B42
         SkinOver        =   "frmMain.frx":458DC
         SkinUp          =   "frmMain.frx":49676
         MaskColor       =   8421504
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdBar 
         Height          =   1050
         Index           =   6
         Left            =   270
         TabIndex        =   10
         Tag             =   "Sales"
         Top             =   1230
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Appearance      =   4
         AutoMask        =   0   'False
         BackColor       =   -2147483636
         Caption         =   ""
         CaptionOffsetY  =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   8421504
         SkinDisabled    =   "frmMain.frx":4D410
         SkinOver        =   "frmMain.frx":511AA
         SkinUp          =   "frmMain.frx":54F44
         MaskColor       =   8421504
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdBar 
         Height          =   1050
         Index           =   7
         Left            =   270
         TabIndex        =   11
         Tag             =   "Stock"
         Top             =   5580
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Appearance      =   4
         AutoMask        =   0   'False
         BackColor       =   -2147483636
         Caption         =   ""
         CaptionOffsetY  =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   8421504
         SkinDisabled    =   "frmMain.frx":58CDE
         SkinOver        =   "frmMain.frx":5CA78
         SkinUp          =   "frmMain.frx":60812
         MaskColor       =   8421504
         ShowFocus       =   0
      End
      Begin MSForms.Image picSideBar 
         Height          =   9285
         Left            =   60
         Top             =   60
         Width           =   1495
         BackColor       =   8421504
         BorderStyle     =   0
         SpecialEffect   =   2
         Size            =   "2637;16378"
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":645AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65288
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6791C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":685F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":692D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69FB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AC8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B968
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C644
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D31E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73080
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":740D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   5730
   End
   Begin MSComctlLib.StatusBar stbBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   10305
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14049
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/29/2013"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   14055
      TabIndex        =   1
      Top             =   810
      Width           =   14085
      Begin MSForms.Image picBar 
         Height          =   255
         Left            =   -30
         Top             =   0
         Width           =   15165
         BorderStyle     =   0
         SizeMode        =   1
         Size            =   "26749;450"
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   1429
      ButtonWidth     =   1191
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "New"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Inquiry"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Export"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageIndex      =   14
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Adressbook"
         Height          =   645
         Left            =   11940
         Picture         =   "frmMain.frx":743EC
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Addressbook"
         Top             =   120
         Width           =   1005
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   13290
         ScaleHeight     =   525
         ScaleWidth      =   1365
         TabIndex        =   23
         Top             =   180
         Width           =   1365
         Begin MSForms.Label Label1 
            Height          =   345
            Left            =   240
            TabIndex        =   24
            Top             =   180
            Width           =   795
            ForeColor       =   8421504
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "Log off"
            Size            =   "1402;609"
            FontName        =   "Arial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.Label Label2 
            Height          =   345
            Left            =   90
            TabIndex        =   30
            Top             =   135
            Width           =   1275
            ForeColor       =   8421504
            BackColor       =   16777215
            VariousPropertyBits=   8388627
            Caption         =   "8"
            Size            =   "2249;609"
            FontName        =   "Webdings"
            FontHeight      =   285
            FontCharSet     =   2
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   14670
         Picture         =   "frmMain.frx":746F6
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   12
         ToolTipText     =   " Log Off "
         Top             =   90
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog cmLogo 
      Left            =   2100
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save file to..."
      Filter          =   "*.xls"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2100
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save file to..."
      Filter          =   "*.xls"
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPriceChange 
         Caption         =   "Price Change Batch"
      End
      Begin VB.Menu Line10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAppro 
         Caption         =   "Appro's"
      End
      Begin VB.Menu Line9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPastel 
         Caption         =   "Pastel Partner"
         Begin VB.Menu mnuPas2004 
            Caption         =   "Pastel Partner 2004 Interface"
         End
         Begin VB.Menu mnuPas2005 
            Caption         =   "Pastel Partner 2005 Interface"
         End
      End
   End
   Begin VB.Menu mnu1settings 
      Caption         =   "Settings"
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnusubscription 
         Caption         =   "Subscription settings"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDebtJournal 
         Caption         =   "&Pass a Debtors Journal"
      End
      Begin VB.Menu mnuCredJournal 
         Caption         =   "Pass a &Creditors Journal"
      End
      Begin VB.Menu Line11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProduction 
         Caption         =   "Portion Production"
      End
      Begin VB.Menu line45 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReCalc 
         Caption         =   "&Recalculate Recipe Costs"
      End
      Begin VB.Menu mnuReCalcCon 
         Caption         =   "Recalculate &Consumption"
      End
      Begin VB.Menu mnuDebtReCalc 
         Caption         =   "Recalculate &Debtor Balances"
      End
      Begin VB.Menu mnuSuppReCalc 
         Caption         =   "Recalculate &Supplier Balances"
      End
      Begin VB.Menu Line67 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinks 
         Caption         =   "Link Products to Suppliers"
      End
      Begin VB.Menu Line6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKitchen 
         Caption         =   "Kitchen Printing Setup"
      End
      Begin VB.Menu Line7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTelEx 
         Caption         =   "Telephone Exchanges"
         Begin VB.Menu mnu3000 
            Caption         =   "Man 3000"
         End
         Begin VB.Menu Line8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTelTrace 
            Caption         =   "Teltrace SOHO"
         End
      End
      Begin VB.Menu yu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCipher 
         Caption         =   "Cipher Lab 8001-L Configuration"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTable 
         Caption         =   "Clear Locked Tables"
      End
      Begin VB.Menu mnuTabs 
         Caption         =   "Clear Locked Tabs"
      End
      Begin VB.Menu lines 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAve 
         Caption         =   "Set Average Cost Equal to Landed Cost"
      End
      Begin VB.Menu lines2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear System Totals"
      End
   End
   Begin VB.Menu mndb 
      Caption         =   "Database"
      Begin VB.Menu mnubckup 
         Caption         =   "Backup"
         Begin VB.Menu mnubackup 
            Caption         =   "Backup database"
         End
      End
      Begin VB.Menu mndbUp 
         Caption         =   "Database Update"
         Begin VB.Menu mnDataCominst 
            Caption         =   "Install database update utility"
         End
         Begin VB.Menu mnuEditdbset 
            Caption         =   "Edit settings"
         End
         Begin VB.Menu mndbcom 
            Caption         =   "Update database structure"
         End
         Begin VB.Menu mnuValues 
            Caption         =   "Update database values"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PrintTransfer()
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
    For i = 1 To .grdHead.Rows - 1
        If Trim(.grdHead.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdHead.TextMatrix(i, 1))
                Case "Line Feeds"
                    If Slip_Printer_Type = 0 Then
                        Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdHead.TextMatrix(i, 0)));
                    End If
                Case Else
                    Select Case .grdHead.TextMatrix(i, 2)
                        Case "Left"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdHead.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "STOCK TRANSFER"
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    ActiveReadServer "SELECT TOP 100 PERCENT dbo.Transfer_Journal.Product_Code, dbo.Products.Description, dbo.Locations.Loc_Name AS Out_Name," & _
    " Locations_1.Loc_Name AS In_Name, dbo.Transfer_Journal.Trans_Location_No, dbo.Transfer_Journal.Rec_Location_No, dbo.Transfer_Journal.Qty," & _
    " dbo.Transfer_Journal.Ave_Cost , dbo.Transfer_Journal.Transfer_No" & _
    " FROM dbo.Transfer_Journal INNER JOIN" & _
    " dbo.Locations ON dbo.Transfer_Journal.Trans_Location_No = dbo.Locations.Location_No INNER JOIN" & _
    " dbo.Locations Locations_1 ON dbo.Transfer_Journal.Rec_Location_No = Locations_1.Location_No LEFT OUTER JOIN" & _
    " dbo.Products ON dbo.Transfer_Journal.Product_Code = dbo.Products.Product_Code" & _
    " Where (dbo.Transfer_Journal.Transfer_No = " & Val(frmMain.Toolbar1.Tag) & ")" & _
    " ORDER BY dbo.Transfer_Journal.Line_No"
    If rs.EOF = False Then
        Print #filenum, UCase("From: " & rs.Fields("Trans_Location_No") & " - " & rs.Fields("Out_Name"))
        Print #filenum, UCase("  To: " & rs.Fields("Rec_Location_No") & " - " & rs.Fields("In_Name"))
    End If
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    Total = 0
    While Not rs.EOF
        Print #filenum, rs.Fields("Product_Code") & " - " & rs.Fields("Description")
        Print #filenum, rs.Fields("Qty") & " @ R" & Format(rs.Fields("Ave_Cost"), "0.00") & " = R" & Format(rs.Fields("Ave_Cost") * rs.Fields("Qty"), "0.00")
        Total = Total + rs.Fields("Ave_Cost") * rs.Fields("Qty")
        rs.MoveNext
    Wend
    rs.Close
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    'Print #filenum, Chr(27) & Chr(77) & Chr(49);
    'Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, "TOTAL: R" & Format(Total, "0.00")
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, "Date: " & Format(Date, "dd MMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
    Print #filenum, "Transfer No: " & Format(frmMain.Toolbar1.Tag, "000000")
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
Private Sub cmdBar_Click(Index As Integer)
    If frmMain.Command1.Tag = "Addressbook" Then
    MsgBox "Please use the Close button at the bottom of the page"
    Exit Sub
    End If
    picProdBar.Visible = False
    picAccBar.Visible = False
    Select Case Index
        Case 0
            checked = Checksubscriptiondb
            If checked = True Then
            If UserRecord.Reservations = False Then
                cmdBar(Index).BackColor = vbRed
                MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
                cmdBar(Index).BackColor = &H8000000C
                Exit Sub
            End If
            Screen.MousePointer = 11
            frmDetails.Hide
            cmdBar(0).Enabled = False
            frmRes.Show
            Screen.MousePointer = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Guest Reservations')"
            Else
            MsgBox " Your subscription to HeroPOS has expired please inform HeroPOS to recieve a new date key"
            End If
        Case 1
            If UserRecord.Rooms = False Then
                cmdBar(Index).BackColor = vbRed
                MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
                cmdBar(Index).BackColor = &H8000000C
                Exit Sub
            End If
            Screen.MousePointer = 11
            frmDetails.Hide
            cmdBar(1).Enabled = False
            frmRooms.Show
            Screen.MousePointer = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Rooms')"
        Case 2
            If UserRecord.Guests = False Then
                cmdBar(Index).BackColor = vbRed
                MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
                cmdBar(Index).BackColor = &H8000000C
                Exit Sub
            End If
            Screen.MousePointer = 11
            picAccBar.Visible = True
            frmDetails.Hide
            cmdBar(2).Enabled = False
            frmGuests.Show
            Screen.MousePointer = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Guests')"
        Case 3
'            ' mornegrvedit
'            'If UserRecord.Checkin = False And UserRecord.Checkout = False Then
'            If UserRecord.Inventory = False Then
'                cmdBar(Index).BackColor = vbRed
'                MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
'                cmdBar(Index).BackColor = &H8000000C
'                Exit Sub
'            End If
'            Screen.MousePointer = 11
'            frmDetails.Hide
'            cmdBar(3).Enabled = False
'            frmGRVEDIT.Show
'            Screen.MousePointer = 0
'            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Grv edit and Backorders')"
        Case 4
            If UserRecord.Users = False Then
                cmdBar(Index).BackColor = vbRed
                MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
                cmdBar(Index).BackColor = &H8000000C
                Exit Sub
            End If
            Screen.MousePointer = 11
            frmDetails.Hide
            cmdBar(4).Enabled = False
            frmUsers.Show
            Screen.MousePointer = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Users')"
        Case 5
            If UserRecord.Reports = False Then
                cmdBar(Index).BackColor = vbRed
                MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
                cmdBar(Index).BackColor = &H8000000C
                Exit Sub
            End If
            Screen.MousePointer = 11
            frmDetails.Hide
            cmdBar(5).Enabled = False
            frmReports.Show
            Screen.MousePointer = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Reports')"
        Case 6
            checked = Checksubscriptiondb
            If checked = True Then
                    
            Screen.MousePointer = 11
            If UserRecord.Sales = False Then
                cmdBar(Index).BackColor = vbRed
                MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
                cmdBar(Index).BackColor = &H8000000C
                Exit Sub
            End If
            GlobalMode = TillMode.StartMode
            TillData.Account_No = ""
            TillData.DocNo = 0
            TillData.Res_No = 0
            TillData.Room_No = 0
            TillData.SaleTotal = 0
            TillData.CalculatedTax = 0
            TillData.CollectedTax = 0
            TillData.TableNo = 0
            TillData.Table_Name = ""
            TillData.TabName = ""
            TillData.TabNo = 0
            TillData.TransNo = 0
            TillData.TaxableSales = 0
            TillData.TaxTotal = 0
            frmDetails.Hide
            Screen.MousePointer = 0
            For i = 0 To 7
                If cmdBar(i).Enabled = False Then
                    frmSales.Tag = cmdBar(i).Tag
                End If
            Next i
            cmdBar(6).Enabled = False
            frmMain.Hide
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Sales')"
            
            DoEvents
            frmSales.Show
            Else
            MsgBox " Your subscription to HeroPOS has expired please inform HeroPOS to recieve a new date key"
            End If
            GoTo Getout
        Case 7
            If UserRecord.Inventory = False Then
                cmdBar(Index).BackColor = vbRed
                MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
                cmdBar(Index).BackColor = &H8000000C
                Exit Sub
            End If
            checked = Checksubscriptiondb
            If checked = True Then
                Screen.MousePointer = 11
                frmDetails.Hide
                cmdBar(7).Enabled = False
                For i = 0 To cmdMenu.Count - 1
                    If i = 0 Then
                        cmdMenu(i).FontTextCaption.Bold = True
                        cmdMenu(i).BackColor = &HFFC0C0
                    Else
                        cmdMenu(i).FontTextCaption.Bold = False
                        cmdMenu(i).BackColor = &HFFFFFF
                    End If
                Next i
                DoEvents
                frmProducts.Show
                picProdBar.Visible = True
                Screen.MousePointer = 0
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Inventory')"
            Else
                MsgBox " Your subscription to HeroPOS has expired please inform HeroPOS to recieve a new date key"
            End If
    End Select
    Screen.MousePointer = 11
    For i = 0 To 7
        If i <> Index Then
            If cmdBar(i).Enabled = False Then
                Select Case i
                    Case 0
                        Unload frmRes
                    Case 1
                        Unload frmRooms
                    Case 2
                        Unload frmGuests
                    Case 3
                        Unload frmCheck
                    Case 4
                        If Trim(frmUsers.txtLogin.Text) = "New User" Then
                            frmUsers.txtLogin.Text = ""
                            ActiveUpdateServer "Delete from Users where User_No = " & frmUsers.lblUserNumber
                        End If
                        frmMain.Toolbar1.Buttons(2).Enabled = False
                        Unload frmUsers
                    Case 5
                        Unload frmReports
                    Case 7
                        DoEvents
                        For ib = 0 To cmdMenu.Count - 1
                            If cmdMenu(ib).FontTextCaption.Bold = True Then
                                Select Case ib
                                    Case 0: Unload frmProducts
                                    Case 1: Unload frmDepartments
                                    Case 2: Unload frmLocations
                                    Case 3: Unload frmSuppliers
                                    Case 4: Unload frmOrder
                                    Case 5: Unload frmGRV
                                    Case 6: Unload frmTransfers
                                    Case 7: Unload frmWideCount
                                    Case 8
                                    Case 13: Unload frmProduction
                                End Select
                            End If
                        Next ib
                End Select
                cmdBar(i).Enabled = True
            End If
        End If
    Next i
    Screen.MousePointer = 0
    If frmSubscription.Visible = True Then frmSubscription.Visible = False
    If Subscriptloaded = True Then Unload frmSubscription: Subscriptloaded = False
Getout:
End Sub
Public Sub cmdMenu_Click(Index As Integer)
    On Error Resume Next
    DoEvents
    Screen.MousePointer = 11
    For i = 0 To cmdMenu.Count - 1
        If cmdMenu(i).FontTextCaption.Bold = True Then
            Select Case i
                Case 0: Unload frmProducts:
                
                frmMain.cmdMenu(4).Tag = ""
                
                frmMain.stbBar.Panels(3) = "Records = 0"
                Case 1: Unload frmDepartments
                
                frmMain.cmdMenu(4).Tag = ""
                
                frmMain.stbBar.Panels(3) = "Records = 0"
                Case 2: Unload frmLocations
                 frmMain.cmdMenu(4).Tag = ""
                
                frmMain.stbBar.Panels(3) = "Records = 0"
                Case 3: Unload frmSuppliers
                
                frmMain.cmdMenu(4).Tag = ""
                
                frmMain.stbBar.Panels(3) = "Records = 0"
                Case 4: Unload frmOrder
                        frmMain.stbBar.Panels(3) = "Records = 0"
                        
                       
                        frmMain.cmdMenu(4).Tag = ""
                        
                
                Case 5: Unload frmGRV
                
                frmMain.cmdMenu(4).Tag = ""
                
                frmMain.stbBar.Panels(3) = "Records = 0"
                Case 6: Unload frmTransfers
                
                 frmMain.cmdMenu(4).Tag = ""
                
                frmMain.stbBar.Panels(3) = "Records = 0"
                Case 7: Unload frmWideCount
                
                frmMain.cmdMenu(4).Tag = ""
                
                frmMain.stbBar.Panels(3) = "Records = 0"
                Case 8:
                
                frmMain.cmdMenu(4).Tag = ""
                
                frmMain.stbBar.Panels(3) = "Records = 0"
                Case 13: Unload frmProduction
                
                 frmMain.cmdMenu(4).Tag = ""
               
                frmMain.stbBar.Panels(3) = "Records = 0"
                
            End Select
        End If
    Next i
    DoEvents
    For i = 0 To cmdMenu.Count - 1
        If Index = i Then
            cmdMenu(i).FontTextCaption.Bold = True
            cmdMenu(i).BackColor = &HFFC0C0
        Else
            cmdMenu(i).FontTextCaption.Bold = False
            cmdMenu(i).BackColor = &HFFFFFF
        End If
    Next i
    DoEvents
    Select Case cmdMenu(Index).Caption
        Case "Members"
            
            frmGuests.txtSuppNo.SetFocus
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Members')"
            frmGuests.Label5(2).Caption = "Account Details..."
            
            frmGuests.LoadSuppliers
        Case "Travel Agents"
            
            frmGuests.txtSuppNo.SetFocus
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Travel Agents')"
            frmGuests.LoadSuppliers
        Case "Staff Accounts"
           
            frmGuests.txtSuppNo.SetFocus
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Staff Accounts')"
            frmGuests.LoadSuppliers
        Case "Management Accounts"
           
            frmGuests.txtSuppNo.SetFocus
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Management Accounts')"
            frmGuests.LoadSuppliers
        Case "Debtors"
            
            frmGuests.txtSuppNo.SetFocus
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Debtors')"
            frmGuests.Label5(2).Caption = "Account Details..."
            frmGuests.LoadSuppliers
        Case "Locations"
            
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Locations')"
            frmLocations.Show
        Case "Departments"
            
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Departments')"
            frmDepartments.Show
        Case "Production"
            
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Production')"
            frmProduction.Show
        Case "Products"
           
            'If frmSubscription.Visible = True Then Unload frmSubscription
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Products')"
            
            frmProducts.Load_Products
            frmProducts.Show
            
            frmMain.Toolbar1.Buttons(11).Enabled = True
            frmMain.Toolbar1.Buttons(12).Enabled = True
        Case "Orders"
'            If frmBackorders.Visible = True Then Unload frmBackorders
'            If frmSubscription.Visible = True Then Unload frmSubscription
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Purchase Order')"
            frmOrder.Show
        Case "Goods Receiving"
            'If frmBackorders.Visible = True Then Unload frmBackorders
            If frmSubscription.Visible = True Then Unload frmSubscription
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Goods Receiving')"
            frmGRV.Show
        Case "Stock Takes"
'            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'A4 Stock Takes')"
'            frmWideCount.Show
        Case "Transfers"
            'If frmBackorders.Visible = True Then Unload frmBackorders
            If frmSubscription.Visible = True Then Unload frmSubscription
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Stock Transfers')"
            frmTransfers.Show
        Case "Suppliers"
            'If frmBackorders.Visible = True Then Unload frmBackorders
            If frmSubscription.Visible = True Then Unload frmSubscription
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Suppliers')"
            frmSuppliers.Show
            
        Case "Email"
            'If frmBackorders.Visible = True Then Unload frmBackorders
            If frmSubscription.Visible = True Then Unload frmSubscription
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Email')"
            frmEmails.Show
            
            
            
    End Select
    Screen.MousePointer = 0
    On Error GoTo 0
End Sub



Private Sub Command1_Click()
Dim i As Integer
 For Each Form In Forms
                    If Form.Name = "frmMain" Then GoTo skip
                    If Form.Name = "frmAdrbook" Then GoTo skip
                    Unload Form
skip:
                Next Form
                For i = 0 To 7
cmdBar(i).Enabled = True
Next i
On Error GoTo Errorsa
ActiveReadServer4 " Select * from Adressbook ORDER by Last_Name"
'Load FrmAdrbook
Dim Selected2 As String
If rs4.RecordCount > 0 Then
Selected2 = rs4.Fields("ID_NO")
End If
With FrmAdrbook
FrmAdrbook.top = 0
FrmAdrbook.Text1.Text = ""
FrmAdrbook.Text2.Text = ""
FrmAdrbook.Text3.Text = ""
FrmAdrbook.Text4.Text = ""
FrmAdrbook.Text5.Text = ""
FrmAdrbook.Text6.Text = ""
FrmAdrbook.Text7.Text = ""
FrmAdrbook.Text8.Text = ""
FrmAdrbook.Text9.Text = ""
FrmAdrbook.Text10.Text = ""
FrmAdrbook.Text11.Text = ""
FrmAdrbook.Text12.Text = ""
.Text13.Text = ""
.Text14.Text = ""
FrmAdrbook.Label16.Caption = ""
If .Grdadress.Rows = 0 Then Exit Sub

Selected2 = .Grdadress.TextMatrix(.Grdadress.Row, 2)
ActiveReadServer4 " Select * from Adressbook where Id_No = '" & Selected2 & "' ORDER by Last_Name, First_Name"
If rs4.RecordCount > 0 Then
.Text1.Text = rs4.Fields("Id_No")
.Text2.Text = rs4.Fields("First_Name")
.Text3.Text = rs4.Fields("Last_Name")
.Text4.Text = rs4.Fields("Email")
.Text5.Text = rs4.Fields("Tel_No")
.Text6.Text = rs4.Fields("Fax_No")
.Text7.Text = rs4.Fields("Cell_No")
.Text8.Text = rs4.Fields("Address")
.Text9.Text = rs4.Fields("City")
.Text10.Text = rs4.Fields("Post_code")
.Text11.Text = rs4.Fields("Province")
.Text12.Text = rs4.Fields("Country")
.Text13.Text = rs4.Fields("Remarks")
.Text14.Text = rs4.Fields("Title")
.Label16.Caption = Format(rs4.Fields("Last_Booking"), "dd/mm/yyyy")

End If
If rs4.State = 1 Then rs4.Close
End With
FrmAdrbook.top = 0
frmMain.Command1.Tag = "Addressbook"
Errorsa:
End Sub

Private Sub MDIForm_Activate()
    chkdateagain = ""
    picBar.Width = picTop.Width
    picBar.Height = picTop.Height
    picSideBar.Height = picSide.Height - 75
    ActiveReadServer "Select Loc_Name from Locations where Location_No = " & Location_No
    If rs.RecordCount > 0 Then
        Location_Name = rs.Fields("Loc_Name") & ""
    End If
    rs.Close
    stbBar.Panels(3).Text = Workstation_No & " - " & Trim(Workstation_Name) & "                SERVER: " & UCase(Trim(Server.SQL_Name)) & "              DATABASE: " & UCase(Trim(Server.SQL_Database)) & "               LOCATION: " & Location_No & " - " & UCase(Location_Name)
End Sub
Private Sub MDIForm_Load()
       frmMain.Caption = Trim(gblApp_Name)
    For i = 0 To cmdBar.Count - 1
        cmdBar(i).Enabled = True
    Next i
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
    For Each Form In Forms
        Unload Form
    Next Form
On Error GoTo 0
cnnMain.Close
End
End Sub





Private Sub mnDataCominst_Click()
Dim fso As New FileSystemObject
Dim Path As String
On Error GoTo err:
    Path = App.Path & "\Database Update"
    If fso.FolderExists(Path) = False Then fso.CreateFolder (Path)
    iRetVal = SaveResItemToDisk("Config.xml", "CUSTOM", Path & "\Config.xml")
    iRetVal = SaveResItemToDisk("Database.snpx", "CUSTOM", Path & "\Database.snpx")
    iRetVal = SaveResItemToDisk("xSQL.Licensing.v4.dll", "CUSTOM", Path & "\xSQL.Licensing.v4.dll")
    iRetVal = SaveResItemToDisk("xSQL.Schema.Core.dll", "CUSTOM", Path & "\xSQL.Schema.Core.dll")
    iRetVal = SaveResItemToDisk("xSQL.Schema.SqlServer.dll", "CUSTOM", Path & "\xSQL.Schema.SqlServer.dll")
    iRetVal = SaveResItemToDisk("xSQL.SchemaCompare.SqlServer.dll", "CUSTOM", Path & "\xSQL.SchemaCompare.SqlServer.dll")
    iRetVal = SaveResItemToDisk("xSQLSchemaCmd.exe", "CUSTOM", Path & "\xSQLSchemaCmd.exe")
    
   ' Update server and database name in config.xml
   filenum = FreeFile
   Open Path & "\Config.xml" For Input As #filenum
   Dim outtext As String
   Dim intext As String
   Do While Not EOF(filenum)
        Line Input #filenum, intext
        If InStr(1, intext, "DatabaseName") > 0 Then
            outtext = outtext & Replace(intext, "hero_original", Trim(Server.SQL_Database)) & Chr(13)
            Else
            If InStr(1, intext, "CompareSchema") > 0 Then
                outtext = outtext & Replace(intext, "false", "true") & Chr(13)
            Else
                outtext = outtext & intext & Chr(13)
            End If
        End If
   Loop
   Close #filenum
   
   filenum = FreeFile
   Open Path & "\Config.xml" For Output As filenum
   Print #filenum, outtext
   Close #filenum
   
   
   
   
    MsgBox "Database update utility installed to: " & Path
    Exit Sub
err:
    MsgBox "Error saving database update application to " & Path, vbCritical
End Sub
Private Sub mnuEditdbset_Click()

Path = App.Path & "\Database Update"
Shell "notepad.exe " & Path & "\Config.xml", vbNormalFocus
End Sub

Private Sub mndbcom_Click()
Dim str_Com As String
    Path = App.Path & "\Database Update"
    Result = MsgBox("Please make sure you have a backup of the current database (" & Trim(Server.SQL_Database) & ") before continuing." & Chr(13) & "Continue?", vbYesNo)
    If Result = 6 Then
        str_Com = Path & "\xSQLSchemaCmd.exe " & Path & "\Config.xml"
        Shell str_Com, vbNormalFocus
    End If
End Sub

Private Sub mnuAppro_Click()
    frmAppro.Show vbModal
End Sub

Private Sub mnuAve_Click()
    frmAveCost.Show vbModal
End Sub

Private Sub mnubackup_Click()
'Connect(server As String, user As String, password As String, integrated As Boolean)
'BackupDatabaseToFile(database As String, filename As String)
'Disconnect
'
'
'
' Server.SQL_Name = GetSetting(appname:=Trim(gblApp_Name), Section:="Server", key:="Server")
'    Server.SQL_Database = GetSetting(appname:=Trim(gblApp_Name), Section:="Server", key:="SQL_Database")
'    Server.SQL_User = GetSetting(appname:=Trim(gblApp_Name), Section:="Server", key:="SQL_User", Default:="sa")
'    server.SQL_Password =
On Error GoTo trap
Screen.MousePointer = 11
Dim fso As New FileSystemObject
If ComputerNames = UCase(Trim(Server.SQL_Name)) Then
If fso.FolderExists("c:\backups") = False Then
    fso.CreateFolder ("c:\backups")
    End If
    End If
Backupdb
Screen.MousePointer = 0
Exit Sub
trap:

Screen.MousePointer = 1
MsgBox "Backup feature not fully configured on Server so no backup was done!", vbCritical, "HeroPOS"
On Error GoTo 0
End Sub

Private Sub mnuCipher_Click()
    On Error GoTo trap
    Kill App.Path & "\Cipher\Data_read.ini"
    DoEvents
    Shell App.Path & "\Cipher\Data_Read.exe", vbNormalFocus
    On Error GoTo 0
    Exit Sub
trap:
    If err.Number = 53 Then
        Resume Next
    End If
    MsgBox "Cipher Lab not Installed or Configured Properly", vbCritical, "HeroPOS"
    On Error GoTo 0
End Sub
Private Sub mnuClear_Click()
    If UserRecord.Total_Clear = False Then
        MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
        Exit Sub
    End If
    Response = MsgBox("Clearing System Totlas will Clear All Cashup's, Purchace History, Journals and Revenue Totals! " & Chr(13) & _
    "Are you certain you want to do this?", vbYesNo, "System Wide Clear!")
    Select Case Response
        Case vbYes
            frmClear.Show vbModal
            End
    End Select
End Sub
Private Sub mnuCredJournal_Click()
    Load frmJournal
    frmJournal.Tag = "Creditor"
    frmJournal.Show vbModal
End Sub
Private Sub mnuDebtJournal_Click()
    Load frmJournal
    frmJournal.Tag = "Debtor"
    frmJournal.Show vbModal
End Sub
Private Sub mnuDebtReCalc_Click()
    frmDebtReCalc.Show vbModal
End Sub



Private Sub mnuKitchen_Click()
    frmKitchen.Show vbModal
End Sub

Private Sub mnuLinks_Click()
Screen.MousePointer = 11
ActiveUpdateServer "Delete from supplier_Links"
DoEvents
ActiveUpdateServer "Insert into supplier_Links Select Supplier_No,Product_Code,'',max(Price_Invoiced),max(Date_Time) from Purchase_Journal where Supplier_No is not Null and Product_Code not in (Select Product_code from Supplier_Links) group by Product_Code,Supplier_No"
DoEvents
Screen.MousePointer = 0
MsgBox "Supplier Links Completed", vbInformation, "HeroPOS"
End Sub

Private Sub mnuPas2004_Click()
    frmPastel.Show vbModal
End Sub



Private Sub mnuPriceChange_Click()
    frmPriceChanges.Show vbModal
End Sub

Private Sub mnuProduction_Click()
    frmProduction.Show
    
End Sub

Private Sub mnuReCalc_Click()
    frmReCalc.Show vbModal
End Sub
Private Sub mnuReCalcCon_Click()
    frmRecalcCon.Show vbModal
End Sub

Private Sub mnuresdb_Click()

End Sub

Private Sub mnuSettings_Click()
    If UserRecord.Settings = False Then
        MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
        Exit Sub
    End If
    frmSettings.Show vbModal
End Sub

Private Sub mnusubscription_Click()
If frmDetails.Visible = True Then Unload frmDetails
frmSubscription.Show
End Sub

Private Sub mnuSuppReCalc_Click()
    frmSuppReCalc.Show vbModal
End Sub

Private Sub mnuTable_Click()
    Load frmLocked
    frmLocked.Tag = "Tables"
    DoEvents
    frmLocked.Show vbModal
End Sub

Private Sub mnuTabs_Click()
    Load frmLocked
    frmLocked.Tag = "Tabs"
    DoEvents
    frmLocked.Show vbModal
End Sub

Private Sub mnuTelTrace_Click()
    frmTelTrace.Show vbModal
End Sub

Private Sub mnuValues_Click()
updatedatabase
MsgBox "Done", vbInformation

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Subscriptloaded = True Then
    If frmSubscription.Visible = True Then frmSubscription.Visible = False
    End If
    KeyCode = 0
    frmMain.picProdBar.Visible = False
    frmMain.picAccBar.Visible = False
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
    frmSplash.Show
    DoEvents
    If cmdBar(7).Enabled = False Then
        For i = 0 To frmMain.cmdMenu.Count - 1
            Select Case i
                Case 5
                    If frmMain.cmdMenu(5).Value = 1 Then
                        Unload frmGRV
                        Unload rptGRV
                    End If
            End Select
        Next i
    End If
    If cmdBar(2).Enabled = False Then
        Unload frmGuests
    End If
    DoEvents
    frmMain.Hide
    UserRecord.Service_Charge = False
    For ib = 0 To cmdMenu.Count - 1
        If cmdMenu(ib).FontTextCaption.Bold = True Then
            Select Case ib
                Case 0: Unload frmProducts
                Case 1: Unload frmDepartments
                Case 2: Unload frmLocations
                Case 3: Unload frmSuppliers
                Case 4: Unload frmOrder
                Case 5: Unload frmGRV
                Case 6: Unload frmTransfers
                Case 7: Unload frmWideCount
                Case 8
                Case 13: Unload frmProduction
            End Select
        End If
    Next ib


End Sub
Private Sub Timer1_Timer()
    stbBar.Panels(4).Text = Format(Time, "hh:mm:ss")
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    DoEvents
    Select Case Button
        Case "Produce"
            frmProduction.Produce
        Case "Cut"
            Clipboard.Clear
            Clipboard.SetText ActiveForm.ActiveControl.Text
            ActiveForm.ActiveControl.Text = ""
        Case "Copy"
            Clipboard.Clear
            Clipboard.SetText ActiveForm.ActiveControl.Text
        Case "Paste"
            ActiveForm.ActiveControl.Text = Clipboard.GetText()
            Clipboard.Clear
        Case "Place"
            frmOrder.Accept_Order
            If frmMain.Toolbar1.Tag <> "" Then
                rptOrder.Show
                DoEvents
                frmMain.Toolbar1.Tag = ""
            End If
        Case "End-of-Day"
            Screen.MousePointer = 11
            frmReports.Tag = "1"
            If TradePrint = 1 Then
                Load rptTradeSum
                EOD_Slip
            Else
            
                rptTradeSum.PrintReport False
                DoEvents
                Unload rptTradeSum
                If RAPrint = 1 Then
                    frmReports.Tag = "1"
                    rptReceipts.PrintReport False
                    DoEvents
                    Unload rptReceipts
                End If
                If ChargePrint = 1 Then
                    With frmReports
                        .cmb2.Clear
                        .cmb2.AddItem "Purchase Journal"
                        .cmb2.AddItem "Sales Journal"
                        .cmb2.AddItem "Transfer Journal"
                        .cmb2.AddItem "Payout Journal"
                        .cmb2.AddItem "User Journals"
                        .cmb2.AddItem "Sales Corrections"
                        .cmb2.AddItem "Discounted Sales"
                        .cmb2.Text = "Sales Journal"
                        .cmb3.Text = "Charge Sales"
                    End With
                    frmReports.Selection_Change
                    DoEvents
                    frmReports.Tag = "1"
                    SaveRecords6
                    rptSalesJournal.PrintReport False
                    DoEvents
                    Unload rptSalesJournal
                    With frmReports
                        .cmb2.AddItem "Trade Analysis"
                        '.cmb2.AddItem "Sales Analysis (Graph)"
                        .cmb2.AddItem "Sales Analysis by Product"
                        .cmb2.AddItem "Sales Analysis by Department"
                        .cmb2.AddItem "Sales Analysis by Location"
                        .cmb2.AddItem "Sales Analysis by Debtor"
                        .cmb2.AddItem "Sales Analysis by User"
                        .cmb2.AddItem "Sales Analysis by Supplier"
                        .cmb2.AddItem "Sales Analysis by Hour"
                        .cmb2.AddItem "Product Analysis"
                        If Branch_Type = 10 Then .cmb2.AddItem "Pre-Sales Analysis by Price"
                        .cmb2.Text = "Trade Analysis"
                    End With
                End If
            End If
            If PayoutPrint = 1 Then
                With frmReports
                    .cmb2.Clear
                    .cmb2.AddItem "Purchase Journal"
                    .cmb2.AddItem "Sales Journal"
                    .cmb2.AddItem "Transfer Journal"
                    .cmb2.AddItem "Payout Journal"
                    .cmb2.AddItem "User Journals"
                    .cmb2.AddItem "Sales Corrections"
                    .cmb2.AddItem "Discounted Sales"
                    .cmb2.Text = "Payout Journal"
                End With
                frmReports.Selection_Change
                DoEvents
                frmReports.Tag = "1"
                SaveRecords4
                rptPayJournal.PrintReport False
                DoEvents
                Unload rptPayJournal
                With frmReports
                    .cmb2.AddItem "Trade Analysis"
''                    .cmb2.AddItem "Sales Analysis (Graph)"
                    .cmb2.AddItem "Sales Analysis by Product"
                    .cmb2.AddItem "Sales Analysis by Department"
                    .cmb2.AddItem "Sales Analysis by Location"
                    .cmb2.AddItem "Sales Analysis by Debtor"
                    .cmb2.AddItem "Sales Analysis by User"
                    .cmb2.AddItem "Sales Analysis by Supplier"
                    .cmb2.AddItem "Sales Analysis by Hour"
                    .cmb2.AddItem "Product Analysis"
                    If Branch_Type = 10 Then .cmb2.AddItem "Pre-Sales Analysis by Price"
                    .cmb2.Text = "Trade Analysis"
                End With
            End If
            Screen.MousePointer = 0
        Case "Export"
            cmLogo.InitDir = App.Path & "\Export"
            cmLogo.Action = 2
            DoEvents
            If cmLogo.FileName <> "" Then
                FileName = cmLogo.FileName
                If Right(FileName, 4) <> ".xls" Then FileName = FileName & ".xls"
                Select Case frmMain.Toolbar1.Buttons(16).Tag
                    Case "Products"
                        frmProducts.grdProd.FixedRows = 0
                        frmProducts.grdProd.SaveGrid FileName, flexFileExcel
                        frmProducts.grdProd.FixedRows = 1
                    Case "Locations"
                        frmLocations.grdLoc.FixedRows = 0
                        frmLocations.grdLoc.SaveGrid FileName, flexFileExcel
                        frmLocations.grdLoc.FixedRows = 1
                    Case "Users"
                        frmUsers.grdUsers.FixedRows = 0
                        frmUsers.grdUsers.SaveGrid FileName, flexFileExcel
                        frmUsers.grdUsers.FixedRows = 1
                    Case "Debtors"
                        frmGuests.grdSupp.FixedRows = 0
                        frmGuests.grdSupp.SaveGrid FileName, flexFileExcel
                        frmGuests.grdSupp.FixedRows = 1
                    Case "Check"
                        frmCheck.grdGuest.FixedRows = 0
                        frmCheck.grdGuest.SaveGrid FileName, flexFileExcel
                        frmCheck.grdGuest.FixedRows = 1
                    Case "Rooms"
                        frmRooms.grdRooms.FixedRows = 0
                        frmRooms.grdRooms.SaveGrid FileName, flexFileExcel
                        frmRooms.grdRooms.FixedRows = 1
                    Case "Suppliers"
                        frmSuppliers.grdSupp.FixedRows = 0
                        frmSuppliers.grdSupp.SaveGrid FileName, flexFileExcel
                        frmSuppliers.grdSupp.FixedRows = 1
                    Case "Reports"
                        frmReports.grdMain.FixedRows = 0
                        frmReports.grdMain.SaveGrid FileName, flexFileExcel
                        frmReports.grdMain.FixedRows = 1
                End Select
                MsgBox "File saved as " & FileName, vbInformation, "HeroPOS"
            Else
                MsgBox "No Filename specified to save to."
            End If
        Case "Accept"
        
        x = frmMain.picBar.Tag
        
        'Transferfix
        
        
        If frmMain.picBar.Tag = "Transfer" Then
            If frmMain.cmdMenu(6).Value = 1 Then
                frmTransfers.Accept_Transfer
                If frmMain.Toolbar1.Tag <> "" Then
                    If PrintSlipTransfers = 0 Then
                        If Branch_Type = 10 Then
                            rptTransfersW.Show
                        Else
                            rptTransfers.Show
                        End If
                    Else
                        PrintTransfer
                    End If
                    DoEvents
                    frmMain.picBar.Tag = ""
                End If
            End If
            Exit Sub
            End If
            
            If frmGRV.cmbSuppliers.Text = "" Then MsgBox "Please select a supplier first before accepting!": Exit Sub
            If frmMain.cmdMenu(5).Value = 1 Then
                frmGRV.Accept_GRV
                If frmMain.Toolbar1.Tag <> "" Then
                    rptGRV.Show
                    DoEvents
                    frmMain.Toolbar1.Tag = ""
                End If
            End If
        Case "Open"
            If cmdBar(7).Enabled = False Then
                For i = 0 To frmMain.cmdMenu.Count - 1
                    Select Case i
                        Case 4
                            If frmMain.cmdMenu(4).Value = 1 Then
                                frmOrder.OpenOrder
                            End If
                        Case 5
                            If frmMain.cmdMenu(5).Value = 1 Then
                                frmGRV.OpenGRV
                            End If
                    End Select
                Next i
            End If
        Case "Inquiry"
            If cmdBar(0).Enabled = False Then
                frmSearchRes.Show vbModal
                On Error GoTo 0
                Exit Sub
            End If
            If cmdBar(5).Enabled = False Then
                If frmReports.cmdMenu(1).Value = 1 Then
                    Load frmInquiry
                    If frmReports.grdMain.Rows = 1 And frmReports.grdMain.Visible = False Then
                        frmInquiry.Tag = ""
                    Else
                        frmInquiry.Tag = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 0)
                    End If
                    frmInquiry.mthViewStart.Value = Date
                    frmInquiry.mthViewEnd.Value = Date
                    frmInquiry.Show vbModal
                    On Error GoTo 0
                    Exit Sub
                End If
            End If
            
            If cmdBar(7).Enabled = False Then
                If cmdMenu(0).Value = 1 Then
                    Load frmInquiry
                    If frmProducts.grdProd.Rows = 1 Then
                        frmInquiry.Tag = ""
                    Else
                        frmInquiry.Tag = frmProducts.grdProd.TextMatrix(frmProducts.grdProd.Row, 0)
                    End If
                    frmInquiry.mthViewStart.Value = Date
                    frmInquiry.mthViewEnd.Value = Date
                    frmInquiry.Show vbModal
                    On Error GoTo 0
                    Exit Sub
                End If
                If cmdMenu(5).Value = 1 Then
                    Load frmInquiry
                    If frmGRV.grdGRV.Rows = 1 Then
                        frmInquiry.Tag = ""
                    Else
                        frmInquiry.Tag = frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0)
                    End If
                    frmInquiry.mthViewStart.Value = Date
                    frmInquiry.mthViewEnd.Value = Date
                    frmInquiry.Show vbModal
                    On Error GoTo 0
                    Exit Sub
                End If
            End If
            If cmdBar(5).Enabled = False Then
                Select Case frmReports.cmb2.Text
                    Case "Sales Analysis by Product", "Sales Analysis by Location", "Sales Analysis by User", "Sales Analysis by Department"
                        Load frmInquiry
                        If frmReports.grdMain.Rows = 1 Then
                            frmInquiry.Tag = ""
                        Else
                            frmInquiry.Tag = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 0)
                        End If
                        frmInquiry.mthViewStart.Value = frmReports.mthViewStart.Value
                        frmInquiry.mthViewEnd.Value = frmReports.mthViewEnd.Value
                        frmInquiry.Show vbModal
                        On Error GoTo 0
                        Exit Sub
                    Case Else
                        frmInquiry.Show vbModal
                        On Error GoTo 0
                        Exit Sub
                End Select
            End If
            frmInquiry.Show vbModal
        Case "Print"
            If cmdBar(7).Enabled = False Then
                If cmdMenu(5).Value = 1 Then
                    rptGRV.PrintReport True
                End If
            End If
            If cmdBar(4).Enabled = False Then
                frmUsers.Tag = "1"
                rptUsers.PrintReport True
            End If
            If cmdBar(7).Enabled = False Then
                If cmdMenu(0).Value = 1 Then
                    rptProducts.PrintReport True
                End If
            End If
            If cmdBar(2).Enabled = False Then
               If Pricelistpreview = "Yes" Then
               frmGuests.Tag = "1"
               rptPricelist.Print True
               End If
            End If
            
            If cmdBar(5).Enabled = False Then
                Select Case frmReports.cmb2.Text
                    Case "Stock Movement (Values)", "Stock Movement (Quantities)"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptMovement.PrintReport True
                        Screen.MousePointer = 0
                    Case "Age Analysis"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords8
                        rptAgeAnalysis.PrintReport True
                        Screen.MousePointer = 0
                    Case "Stock on Hand (Suppliers)"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptStockonHand.PrintReport True
                        Screen.MousePointer = 0
                    Case "Stock on Hand"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptStockonHand.PrintReport True
                        Screen.MousePointer = 0
                    Case "Product Analysis"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords7
                        rptProductSales.PrintReport True
                        Screen.MousePointer = 0
                    Case "Sales Journal"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords6
                        rptSalesJournal.PrintReport True
                        Screen.MousePointer = 0
                    Case "Receive on Account"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptReceipts.PrintReport True
                        Screen.MousePointer = 0
                     Case "Staff Shift Report"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords5
                        rptShifts.PrintReport True
                        Screen.MousePointer = 0
                    Case "Payout Journal"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords4
                        rptPayJournal.PrintReport True
                        Screen.MousePointer = 0
                    Case "Stock Takes"
                        frmReports.Tag = "1"
                        rptStockTakes.PrintReport True
                    Case "Staff Commision Report"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords3
                        rptUserComm.PrintReport True
                        Screen.MousePointer = 0
                    Case "Stock Variance"
                        frmReports.Tag = "1"
                        rptVariance.PrintReport True
                    Case "Purchase Journal"
                        frmReports.Tag = "1"
                        SaveRecords1
                        rptPurJournal.PrintReport True
                    Case "Trade Analysis"
                        frmReports.Tag = "1"
                        If TradePrint = 1 Then
                            Trade_Print_Slip
                        Else
                            rptTradeSum.PrintReport True
                        End If
                    Case "Sales Analysis by Debtor"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.PrintReport True
                    Case "Sales Analysis by Product"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.PrintReport True
                    Case "Sales Analysis by Location"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.PrintReport True
                    Case "Sales Analysis by User"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.PrintReport True
                    Case "Sales Analysis by Department"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.PrintReport True
                        End Select
                        End If
        Case "Preview"
            If cmdBar(4).Enabled = False Then
                frmUsers.Tag = "1"
                rptUsers.Show
            End If
            If cmdBar(5).Enabled = False Then
                Select Case frmReports.cmb2.Text
                    
                    Case "Sales Analysis by User type"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptSalesanalisysbyallusers.Show
                        Screen.MousePointer = 0
                    Case "Stock Movement (Values)", "Stock Movement (Quantities)"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptMovement.Show
                        Screen.MousePointer = 0
                    Case "Stock on Hand (Suppliers)"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptStockonHand.Show
                        Screen.MousePointer = 0
                    Case "Age Analysis"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords8
                        rptAgeAnalysis.Show
                        Screen.MousePointer = 0
                    Case "Stock on Hand"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptStockonHand.Show
                        Screen.MousePointer = 0
                    
                    Case "Stock Levels Low"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        rptStockonHand.Show
                        Screen.MousePointer = 0
                    
                    
                    
                    Case "Product Analysis"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords7
                        rptProductSales.Show
                        Screen.MousePointer = 0
                    Case "Sales Journal"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords6
                        rptSalesJournal.Show
                        Screen.MousePointer = 0
                    Case "Receive on Account"
                        frmReports.Tag = "1"
                        rptReceipts.Show
                    Case "Staff Shift Report"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords5
                        rptShifts.Show
                        Screen.MousePointer = 0
                    Case "Stock Takes"
                        frmReports.Tag = "1"
                        rptStockTakes.Show
                    Case "Payout Journal"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords4
                        rptPayJournal.Show
                        Screen.MousePointer = 0
                    Case "Staff Commision Report"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords3
                        rptUserComm.Show
                        Screen.MousePointer = 0
                    Case "Room Sales"
                        If frmReports.grdMain.Row <> 0 Then
                            frmReports.Tag = "1"
                            TillData.Res_No = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 3)
                            rptInvoice.Show
                        End If
                    Case "Sales Corrections"
                        frmReports.Tag = "1"
                        rptCorrect.Show
                    Case "Discounted Sales"
                        frmReports.Tag = "1"
                        rptDiscount.Show
                    Case "Deposits Paid"
                    
                    Case "Pre-Sales Analysis by Price"
                        Screen.MousePointer = 11
                        frmReports.Tag = "1"
                        SaveRecords2
                        rptPreSales.Show
                        Screen.MousePointer = 0
                    Case "Stock Variance"
                        frmReports.Tag = "1"
                        rptVariance.Show
                    Case "Purchase Journal"
                        frmReports.Tag = "1"
                        SaveRecords1
                        rptPurJournal.Show
                    Case "Trade Analysis"
                        frmReports.Tag = "1"
                        rptTradeSum.Show
                    Case "Daily Trade Analysis"
                        frmReports.Tag = "1"
                        rptTradeSumdaily.Show
                    Case "Sales Analysis by Debtor"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.Show
                    Case "Sales Analysis by Product"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.Show
                    Case "Sales Analysis by Location"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.Show
                    Case "Sales Analysis by User"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.Show
                    Case "Sales Analysis by Department"
                        frmReports.Tag = "1"
                        SaveRecords
                        rptSalesByProd.Show
                End Select
            End If
            
               
            If cmdBar(7).Enabled = False Then
                If cmdMenu(0).Value = 1 Then
                    rptProducts.Show
                End If
                End If
            
            If cmdBar(2).Enabled = False Then
               If Pricelistpreview = "Yes" Then
               frmGuests.Tag = "1"
               rptPricelist.Show True
               End If
            End If
            
    
            
            
        
        
        
        
        Case "Save"
            If cmdBar(1).Enabled = False Then frmRooms.SaveRooms
            If cmdBar(4).Enabled = False Then SaveUsers
            If cmdBar(2).Enabled = False Then SaveDebtor
            DoSave = True
            For i = 0 To cmdBar.Count - 1
                If cmdBar(i).Enabled = False Then
                    DoSave = False
                End If
           Next i
           If DoSave = True Then SaveDetails
           If cmdBar(7).Enabled = False Then
                For i = 0 To frmMain.cmdMenu.Count - 1
                    Select Case i
                        Case 0: If frmMain.cmdMenu(0).Value = 1 Then SaveProduct
                        Case 1: If frmMain.cmdMenu(1).Value = 1 Then SaveDepartment
                        Case 2: If frmMain.cmdMenu(2).Value = 1 Then SaveLocation
                        Case 3: If frmMain.cmdMenu(3).Value = 1 Then SaveSupplier
                        Case 4: If frmMain.cmdMenu(4).Value = 1 Then frmOrder.SaveOrder
                         
                        'If frmMain.cmdMenu(4).Value = 1 Then frmOrder.SaveOrder
                        
                       
'                        If frmMain.cmdMenu(4).Value = 1 And frmbackOrder.Visible Then
'                        frmBackorders.SaveBackOrder
'                        End If
                        Case 5: If frmMain.cmdMenu(5).Value = 1 Then frmGRV.SaveGrv
                        Case 13: If frmMain.cmdMenu(13).Value = 1 Then frmProduction.SaveProduction
                    End Select
                Next i
            End If
        Case "New"
            If cmdBar(4).Enabled = False Then CreateUser
            If cmdBar(1).Enabled = False Then frmRooms.CreateRooms
            If cmdBar(2).Enabled = False Then CreateDebtor
            If cmdBar(7).Enabled = False Then
                For i = 0 To frmMain.cmdMenu.Count - 1
                    Select Case i
                    
                        Case 0:
                        'If frmMain.cmdMenu(4).Value = 1 Then frmBackorders.CreateNewOrder
                        If frmMain.cmdMenu(0).Value = 1 Then frmProducts.CreateProduct
                        Case 1: If frmMain.cmdMenu(1).Value = 1 Then frmDepartments.CreateDepartment -1
                        Case 2: If frmMain.cmdMenu(2).Value = 1 Then CreateLocation -1
                        Case 3: If frmMain.cmdMenu(3).Value = 1 Then CreateSupplier
                    End Select
                Next i
            End If
        Case "Delete"
            If cmdBar(1).Enabled = False Then
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Delete room')"

                DeleteRoom
            End If
            If cmdBar(2).Enabled = False Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Delete Debtor')"

                DeleteDebtor
            End If
            If cmdBar(5).Enabled = False Then
                If frmReports.cmb2.Text = "Placed Purchase Orders" Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'DELETE Order')"

                DeleteOrder
                End If
            End If
            If cmdBar(7).Enabled = False Then
                If frmMain.cmdMenu(0).Value = 1 Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Deleted a Product')"
                DeleteProduct
                End If
                If frmMain.cmdMenu(2).Value = 1 Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Deleted a Location')"
                DeleteLocation
                End If
                If frmMain.cmdMenu(1).Value = 1 Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Deleted a Department')"
                DeleteDepartment
                End If
                If frmMain.cmdMenu(4).Value = 1 Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Deleted an Order')"
                frmOrder.DeleteOrder
                End If
                If frmMain.cmdMenu(5).Value = 1 Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Deleted a GRV')"
                frmGRV.DeleteGRV
                End If
                If frmMain.cmdMenu(6).Value = 1 Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Deleted a Transfer')"
                frmTransfers.DeleteTransfer
                End If
                If frmMain.cmdMenu(3).Value = 1 Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Deleted a Supplier')"
                DeleteSupplier
                End If
            End If
            
            If cmdBar(4).Enabled = False Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Deleted a User')"
                DeleteUsers
            End If
    End Select
    
    On Error GoTo 0
End Sub
Private Sub DeleteOrder()
    Response = MsgBox("Are you certain you want to delete purchase order no: " & Trim(Mid(frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 3), InStrRev(frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 3), ":") + 1)) & "?", vbYesNo, "HeroPOS")
    If Response = vbYes Then
        ActiveUpdateServer "Delete from Purchase_Order_Journal where Order_No = " & Val(Trim(Mid(frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 3), InStrRev(frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 3), ":") + 1)))
        frmReports.Selection_Change
    End If
End Sub
Private Sub SaveRecords()
    ActiveReadServer1 "Select isnull(max(Report_No),0)+1 as Report_No from Report_Temp"
    frmReports.Report_No = rs1.Fields("Report_No")
    rs1.Close
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO Report_Temp(Product_Code, Description, Qty_Sold, Total_Cost, GP_Percentage, GP_Value, Total_Revenue, Profit_Index, Report_No)" & _
            "VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & .grdMain.ValueMatrix(i, 4) * 100 & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "','" & .grdMain.TextMatrix(i, 7) & "','" & .Report_No & "')"
        Next i
    End With
End Sub
Private Sub SaveRecords1()
    ActiveReadServer1 "Select isnull(max(Report_No),0)+1 as Report_No from PurJournal_Temp"
    frmReports.Report_No = rs1.Fields("Report_No")
    rs1.Close
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO PurJournal_Temp(Date_Time, [User] , Supplier, Document, Transaction_Type, TotalEx, Tax, TotalinC, Report_No)" & _
            "VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & Replace(.grdMain.TextMatrix(i, 4), "%", "") & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "','" & .grdMain.TextMatrix(i, 7) & "','" & .Report_No & "')"
        Next i
    End With
End Sub
Private Sub SaveRecords2()
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO [PreSales_Temp]([Product_Code], [Description], [Landed_Cost], [GP1], [Wholesale], [GP2], [Retail])" & _
            " VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & Replace(.grdMain.TextMatrix(i, 4), "%", "") & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "')"
        Next i
    End With
End Sub
Private Sub SaveRecords3()
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO [Comm_Temp]([User_Name], [Cash], [Card], [Cheque], [Charge], [TotIncl], [Tax],[TotExcl],[Comm],[CommDue])" & _
            " VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & Replace(.grdMain.TextMatrix(i, 4), "%", "") & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "','" & .grdMain.TextMatrix(i, 7) & "','" & .grdMain.TextMatrix(i, 8) & "','" & .grdMain.TextMatrix(i, 9) & "')"
        Next i
    End With
End Sub
Private Sub SaveRecords4()
    ActiveUpdateServer "Delete from Pay_Temp"
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO [Pay_Temp]([Date_Time], [Account], [User_Name], [Description], [Total])" & _
            " VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & .grdMain.TextMatrix(i, 4) & "')"
        Next i
    End With
End Sub
Private Sub SaveRecords5()
    ActiveUpdateServer "Delete from Shift_Temp"
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO [Shift_Temp]([Date_Time], [User_No], [Shift_Start], [Shift_Stop],[Shift_Len])" & _
            " VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & .grdMain.TextMatrix(i, 5) & "')"
        Next i
    End With
End Sub
Private Sub SaveRecords6()
    ActiveUpdateServer "Delete from Sales_Temp"
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO [Sales_Temp]([Date_Time], [User_No], [Detail], [Doc_No],[TransType],[TotalEx],[Tax],[TotalInc])" & _
            " VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & .grdMain.TextMatrix(i, 4) & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "','" & .grdMain.TextMatrix(i, 7) & "')"
        Next i
    End With
End Sub
Private Sub SaveRecords7()
    ActiveUpdateServer "Delete from Prod_Analysis_Temp"
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO [Prod_Analysis_Temp]([Product_Code], [Description], [Department], [Cost], [Markup], [GP], [Selling])" & _
            " VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & .grdMain.TextMatrix(i, 4) & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "')"
        Next i
    End With
End Sub
Private Sub SaveRecords8()
    ActiveUpdateServer "Delete from Age_Temp"
    With frmReports
        For i = 1 To frmReports.grdMain.Rows - 1
            ActiveUpdateServer "INSERT INTO [Age_Temp]([Debt_No], [Debt_Name], [120Days], [90Days], [60Days], [30Days], [Current], [Balance])" & _
            " VALUES('" & .grdMain.TextMatrix(i, 0) & "','" & .grdMain.TextMatrix(i, 1) & "','" & .grdMain.TextMatrix(i, 2) & "','" & .grdMain.TextMatrix(i, 3) & "','" & .grdMain.TextMatrix(i, 4) & "','" & .grdMain.TextMatrix(i, 5) & "','" & .grdMain.TextMatrix(i, 6) & "','" & .grdMain.TextMatrix(i, 7) & "')"
        Next i
    End With
End Sub
Public Sub SaveSupplier()
With frmSuppliers
    For i = 0 To .optTerms.Count - 1
        If .optTerms(i).Value = True Then
            Terms = i
            Exit For
        End If
    Next i
    ActiveReadServer "Select * from Suppliers where Supplier_No = '" & .txtSuppNo.Text & "'"
    If rs.RecordCount = 0 Then
        ActiveUpdateServer "INSERT INTO Suppliers (Supplier_No, Supplier_Name, Contact_Person, Credit_Limit, Business_Tel, Mobile, Fax_Tel, Address, E_Mail, Web_Page, Terms, Balance, VAT_No,GL_Code,Landed_Cost)" & _
        " VALUES('" & .txtSuppNo.Text & "','" & .txtSuppName.Text & "','" & .txtContact.Text & "','" & .txtCredit.Text & "','" & .txtBussTell.Text & "','" & .txtCell.Text & "','" & .txtFax.Text & "','" & .txtAddress.Text & "','" & .txtEmail.Text & "','" & .txtWeb.Text & "'," & Terms & ",0,'" & .txtVat & "','" & .txtGL_Code.Text & "','" & Val(.chkLand.Value & "") & "')"
        .grdSupp.Rows = .grdSupp.Rows + 1
        .grdSupp.Tag = "1"
        .grdSupp.Row = .grdSupp.Rows - 1
        .grdSupp.ShowCell .grdSupp.Row, 0
        .grdSupp.TextMatrix(.grdSupp.Row, 4) = "0.00"
        .grdSupp.Tag = ""
    Else
        ActiveUpdateServer "UPDATE Suppliers" & _
        " SET " & _
        " Supplier_Name='" & .txtSuppName.Text & "'," & _
        " Contact_Person='" & .txtContact.Text & "'," & _
        " Credit_Limit=" & .txtCredit & "," & _
        " Business_Tel='" & .txtBussTell.Text & "'," & _
        " Mobile='" & .txtCell.Text & "'," & _
        " Fax_Tel='" & .txtFax.Text & "'," & _
        " Address='" & .txtAddress.Text & "'," & _
        " E_Mail='" & .txtEmail.Text & "'," & _
        " Web_Page='" & .txtWeb.Text & "'," & _
        " VAT_No ='" & .txtVat.Text & "'," & _
        " GL_Code ='" & .txtGL_Code.Text & "'," & _
        " Landed_Cost ='" & Val(.chkLand.Value & "") & "'," & _
        " Terms=" & i & _
        " Where Supplier_No = '" & .txtSuppNo.Text & "'"
        Val (.chkLand.Value & "")
    End If
    rs.Close
    .grdSupp.TextMatrix(.grdSupp.Row, 0) = .txtSuppNo.Text
    .grdSupp.TextMatrix(.grdSupp.Row, 1) = .txtSuppName.Text
    .grdSupp.TextMatrix(.grdSupp.Row, 2) = .txtContact.Text
    .grdSupp.TextMatrix(.grdSupp.Row, 3) = .txtBussTell.Text
    MsgBox "Supplier update successful", vbInformation, "HeroPOS"
    frmMain.Toolbar1.Buttons(4).Enabled = False
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(5).Enabled = True
End With
End Sub
Public Sub SaveDebtor()
With frmGuests
    For i = 0 To .optTerms.Count - 1
        If .optTerms(i).Value = True Then
            Terms = i
            Exit For
        End If
    Next i
    If cmdMenu(11).Value <> 0 Then
        DType = 0
    End If
    If cmdMenu(9).Value <> 0 Then
        DType = 1
    End If
    If cmdMenu(10).Value <> 0 Then
        DType = 2
    End If
    If cmdMenu(12).Value <> 0 Then
        DType = 3
    End If
    If cmdMenu(14).Value <> 0 Then
        DType = 4
    End If
    ActiveReadServer "Select * from Debtors where Debtor_No = '" & .txtSuppNo.Text & "'"
    If rs.RecordCount = 0 Then
        ActiveUpdateServer "INSERT INTO Debtors (Debtor_No, Debtor_Name, Contact_Person, Credit_Limit, Business_Tel, Mobile, Fax_Tel, Address, E_Mail, Web_Page, Terms, Balance, VAT_No,GL_Code,Debt_Type)" & _
        " VALUES('" & .txtSuppNo.Text & "','" & .txtSuppName.Text & "','" & .txtContact.Text & "','" & .txtCredit.Text & "','" & .txtBussTell.Text & "','" & .txtCell.Text & "','" & .txtFax.Text & "','" & .txtAddress.Text & "','" & .txtEmail.Text & "','" & .txtWeb.Text & "'," & Terms & ",0,'" & .txtVat & "','" & .txtGL_Code.Text & "'," & DType & ")"
        .grdSupp.Rows = .grdSupp.Rows + 1
        .grdSupp.Tag = "1"
        .grdSupp.Row = .grdSupp.Rows - 1
        .grdSupp.ShowCell .grdSupp.Row, 0
        .grdSupp.TextMatrix(.grdSupp.Row, 4) = "0.00"
        .grdSupp.Tag = ""
    Else
        ActiveUpdateServer "UPDATE Debtors" & _
        " SET " & _
        " Debtor_Name='" & .txtSuppName.Text & "'," & _
        " Contact_Person='" & .txtContact.Text & "'," & _
        " Credit_Limit=" & .txtCredit & "," & _
        " Business_Tel='" & .txtBussTell.Text & "'," & _
        " Mobile='" & .txtCell.Text & "'," & _
        " Fax_Tel='" & .txtFax.Text & "'," & _
        " Address='" & .txtAddress.Text & "'," & _
        " E_Mail='" & .txtEmail.Text & "'," & _
        " Web_Page='" & .txtWeb.Text & "'," & _
        " VAT_No ='" & .txtVat.Text & "'," & _
        " GL_Code ='" & .txtGL_Code.Text & "'," & _
        " Debt_Type =" & DType & "," & _
        " Terms=" & i & _
        " Where Debtor_No = '" & .txtSuppNo.Text & "'"
    End If
    rs.Close
    .grdSupp.TextMatrix(.grdSupp.Row, 0) = .txtSuppNo.Text
    .grdSupp.TextMatrix(.grdSupp.Row, 1) = .txtSuppName.Text
    .grdSupp.TextMatrix(.grdSupp.Row, 2) = .txtContact.Text
    .grdSupp.TextMatrix(.grdSupp.Row, 3) = .txtBussTell.Text
    MsgBox "Debtor update successful", vbInformation, "HeroPOS"
    frmMain.Toolbar1.Buttons(4).Enabled = False
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(5).Enabled = True
End With
End Sub
Public Sub CreateSupplier()
    With frmSuppliers
        .txtSuppNo = ""
        .txtSuppName = ""
        .txtContact = ""
        .txtAddress = ""
        .txtBussTell = ""
        .txtCell = ""
        .txtEmail = ""
        .txtCredit = "0.00"
        .txtFax = ""
        .txtVat = ""
        .txtWeb = ""
        .txtGL_Code.Text = ""
        .optTerms(1).Value = True
        frmMain.Toolbar1.Buttons(2).Enabled = False
        .txtSuppNo.SetFocus
    End With
End Sub
Public Sub CreateDebtor()
    With frmGuests
        .txtSuppNo = ""
        .txtSuppName = ""
        .txtContact = ""
        .txtAddress = ""
        .txtBussTell = ""
        .txtCell = ""
        .txtEmail = ""
        .txtCredit = "0.00"
        .txtFax = ""
        .txtVat = ""
        .txtWeb = ""
        .txtGL_Code.Text = ""
        .optTerms(1).Value = True
        frmMain.Toolbar1.Buttons(2).Enabled = False
        .txtSuppNo.SetFocus
    End With
End Sub
Public Sub SaveUsers()
On Error Resume Next
With frmUsers
    Select Case .optPassExp(0).Value
        Case False: expire = 0
        Case True: expire = 1
    End Select
    Select Case .cmbGender.Text
        Case "Male": gender = 0
        Case "Female": gender = 1
    End Select
    Select Case .cmbUserType.Text
    Case "Manager": Vtype = 0
    Case "Night Manager": Vtype = 1
    Case "Reservations Clerk": Vtype = 2
    Case "Waiter": Vtype = 3
    Case "Barman": Vtype = 4
    Case "GRV Clerk": Vtype = 5
    Case "Buyer": Vtype = 6
    Case "Supervisor": Vtype = 7
    Case "Cashier": Vtype = 8
    Case "Owner": Vtype = 9
    Case "Staff Member": Vtype = 10
    End Select
    ActiveUpdateServer "Update users set " & _
            "User_Name = '" & .txtLogin.Text & "'" & _
            ",User_Password='" & .txtPassword.Text & "'" & _
            ",First_Name='" & .txtFirstName.Text & "'" & _
            ",Last_Name='" & .txtLastName.Text & "'" & _
            ",Expires=" & expire & _
            ",Gender=" & gender & _
            ",User_Type=" & Vtype & _
            " where user_no =" & .lblUserNumber.Caption
    MsgBox "User update successful", vbInformation
    SaveRow = .grdUsers.Row
    .grdUsers.Rows = 1
    ActiveReadServer "Select User_no, User_Name, First_Name, Last_Name, isnull(Gender,0) as Gender, isnull(User_Type,0) as User_Type From Users order by User_No"
    While Not rs.EOF
            .grdUsers.Rows = .grdUsers.Rows + 1
            .grdUsers.TextMatrix(.grdUsers.Rows - 1, 0) = rs.Fields("User_No")
            .grdUsers.TextMatrix(.grdUsers.Rows - 1, 1) = rs.Fields("User_Name") & ""
            .grdUsers.TextMatrix(.grdUsers.Rows - 1, 2) = rs.Fields("First_Name") & " " & rs.Fields("Last_Name") & ""
            Select Case rs.Fields("Gender")
                Case False: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 3) = "Male"
                Case True: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 3) = "Female"
            End Select
            Select Case rs.Fields("User_Type")
                Case 0: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Manager"
                Case 1: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Night Manager"
                Case 2: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Reservations Clerk"
                Case 3: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Waiter"
                Case 4: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Barman"
                Case 5: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "GRV Clerk"
                Case 6: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Buyer"
                Case 7: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Supervisor"
                Case 8: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Cashier"
                Case 9: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Owner"
                Case 10: .grdUsers.TextMatrix(.grdUsers.Rows - 1, 4) = "Staff Member"
            End Select
        rs.MoveNext
    Wend
    rs.Close
    .grdUsers.Row = SaveRow
      frmMain.Toolbar1.Buttons(2).Enabled = True
      frmUsers.btnSetting.Enabled = True
    On Error GoTo 0
End With
End Sub
Public Sub SaveDepartment()
With frmDepartments
    Select Case .lblType.Caption
        Case "Major Department"
            TaxRate = Mid(.cmbTax1.Text, InStr(.cmbTax1.Text, "-") + 2, InStr(.cmbTax1.Text, "%") - InStr(.cmbTax1.Text, "-") - 2)
            TaxType = Mid(.cmbTax1.Text, 1, InStr(.cmbTax1.Text, "-") - 2)
            If .grdMajor.FindRow(.txtDep.Text, 0, 0, 1, 1) = -1 Then
                .grdMajor.Rows = .grdMajor.Rows + 1
                .grdMajor.TextMatrix(.grdMajor.Rows - 1, 0) = .txtDep.Text
                .grdMajor.TextMatrix(.grdMajor.Rows - 1, 1) = .txtDepName.Text
            Else
                .grdMajor.TextMatrix(.grdMajor.Row, 0) = .txtDep.Text
                .grdMajor.TextMatrix(.grdMajor.Row, 1) = .txtDepName.Text
            End If
            
            Vtype = 0
            pTag = "0"
            Query = "IF not exists (select Department_no from [Departments] where Department_no = '" & .txtDep.Text & "')" & _
                 " BEGIN " & _
                    "INSERT INTO [Departments]([Department_No], [Dept_Name], [Short_Name], [Dept_Type], [Dept_Parent],Range_Start,Range_Stop,Sales_Tax,Tax_Type)" & _
                    "VALUES('" & .txtDep.Text & "','" & .txtDepName.Text & "','" & .txtShortName.Text & "'," & Vtype & ",'" & pTag & "','" & .txtPluStart.Text & "','" & .txtPluStop.Text & "'," & TaxRate & "," & TaxType & ")" & _
                " End" & _
                " Else" & _
                " BEGIN " & _
                    "Update [Departments] " & _
                    "SET " & _
                    "[Dept_Name]='" & .txtDepName.Text & "'," & _
                    "[Short_Name]='" & .txtShortName.Text & "'," & _
                    "[Dept_Type]=" & Vtype & "," & _
                    "[Sales_Tax]=" & TaxRate & "," & _
                    "[Tax_Type]=" & TaxType & "," & _
                    "[Range_Start]='" & .txtPluStart.Text & "'," & _
                    "[Range_Stop]='" & .txtPluStop.Text & "'," & _
                    "[Dept_Parent]='" & pTag & "'" & _
                    " where Department_no = '" & .txtDep.Text & "'" & _
                " End"
            ActiveUpdateServer Query
            ActiveUpdateServer "Delete from Department_Links where Dept_No='" & .txtDep.Text & "'"
            For i = 1 To .grdMinor.Rows - 1
                If .grdMinor.ValueMatrix(i, 2) = True Then
                    Query = "Insert into Department_Links (Dept_No,Location_No) values ('" & .txtDep.Text & "'," & .grdMinor.TextMatrix(i, 0) & " )"
                    ActiveUpdateServer Query
                End If
            Next i

        Case "Sub Department"
            Vtype = 1
            pTag = .grdMajor.TextMatrix(.grdMajor.Row, 0)
            
            If InStr(.txtDep.Text, "-") = 0 Then
                .grdSub.Rows = .grdSub.Rows + 1
                .grdSub.TextMatrix(.grdSub.Rows - 1, 0) = .grdMajor.TextMatrix(.grdMajor.Row, 0) & "-" & .txtDep.Text
                .grdSub.TextMatrix(.grdSub.Rows - 1, 1) = .txtDepName.Text
                DepartmentNo = .grdMajor.TextMatrix(.grdMajor.Row, 0) & "-" & .txtDep.Text
            Else
                DepartmentNo = .txtDep.Text
                .grdSub.TextMatrix(.grdSub.Row, 0) = .txtDep.Text
                .grdSub.TextMatrix(.grdSub.Row, 1) = .txtDepName.Text
            End If
            
             
            Query = "IF not exists (select Department_no from [Departments] where Department_no = '" & DepartmentNo & "')" & _
                 " BEGIN " & _
                    "INSERT INTO [Departments]([Department_No], [Dept_Name], [Short_Name], [Dept_Type], [Dept_Parent],[Zero_Price], [Report_forreign_curr])" & _
                    "VALUES('" & DepartmentNo & "','" & .txtDepName.Text & "','" & .txtShortName.Text & "'," & Vtype & ",'" & pTag & "','" & Abs(.chkZero.Value) & "','" & .chk_foreign_currency.Value & "')" & _
                " End" & _
                " Else" & _
                " BEGIN " & _
                    "Update [Departments] " & _
                    "SET " & _
                    "[Dept_Name]='" & .txtDepName.Text & "'," & _
                    "[Short_Name]='" & .txtShortName.Text & "'," & _
                    "[Dept_Type]=" & Vtype & "," & _
                    "[Dept_Parent]='" & pTag & "'," & _
                    "[Zero_Price]='" & Abs(.chkZero.Value) & "'," & _
                    "[Report_forreign_curr]='" & .chk_foreign_currency.Value & "'" & _
                    " where Department_no = '" & DepartmentNo & "'" & _
                " End"
            ActiveUpdateServer Query
            ActiveUpdateServer "Delete from Department_Links where Dept_No='" & DepartmentNo & "'"
            For i = 1 To .grdMinor.Rows - 1
                If .grdMinor.ValueMatrix(i, 2) = True Then
                    Query = "Insert into Department_Links (Dept_No,Location_No) values ('" & DepartmentNo & "'," & .grdMinor.TextMatrix(i, 0) & " )"
                    ActiveUpdateServer Query
                End If
            Next i

    End Select
    MsgBox "Department update successful", vbInformation
    .txtDep.SetFocus
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(4).Enabled = False
End With
End Sub
Public Sub SaveLocation()
With frmLocations
    Select Case .cmbStockBehave.Text
        Case "Enter Count": bType = 0
        Case "Use System Levels": bType = 1
        Case "Reset to Zero": bType = 2
    End Select
    
    Select Case .cmbLocType.Text
        Case "Sales Location": Vtype = 0
        Case "Stock Location": Vtype = 1
        Case "Expence Location": Vtype = 2
        Case "Outside Location": Vtype = 3
    End Select
    
    ActiveUpdateServer "Update Locations set " & _
            "Loc_Name='" & .txtLocName.Text & "'" & _
            ",Loc_Type=" & Vtype & _
            ",Stock_Take=" & bType & _
            " where Location_no =" & .txtLocNum.Text
            
    ActiveUpdateServer "Delete from Location_Links where Location_No=" & .txtLocNum.Text
    
     ActiveUpdateServer "Insert into Location_Links (Location_No,Panel1,Panel2,Panel3) " & _
     "values (" & .txtLocNum.Text & "," & .chkPanels(0).Tag & "," & .chkPanels(1).Tag & "," & .chkPanels(2).Tag & ")"
    
    MsgBox "Location update successful", vbInformation
    SaveRow = .grdLoc.Row
    .grdLoc.Rows = 1
    ActiveReadServer "Select * From Locations order by Location_No"
    While Not rs.EOF
            .grdLoc.Rows = .grdLoc.Rows + 1
            .grdLoc.TextMatrix(.grdLoc.Rows - 1, 0) = rs.Fields("Location_No")
            .grdLoc.TextMatrix(.grdLoc.Rows - 1, 1) = rs.Fields("Loc_Name") & ""
            Select Case rs.Fields("Loc_Type")
                Case 0: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 2) = "Sales Location"
                Case 1: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 2) = "Stock Location"
                Case 2: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 2) = "Expence Location"
                Case 3: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 2) = "Outside Location"
            End Select
            Select Case rs.Fields("Stock_Take")
                Case 0: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 3) = "Enter Count"
                Case 1: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 3) = "Use System Levels"
                Case 2: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 3) = "Reset to Zero"
            End Select
        rs.MoveNext
    Wend
    rs.Close
    .grdLoc.Row = SaveRow
    frmMain.Toolbar1.Buttons(2).Enabled = True
End With
End Sub
Public Sub SaveProduct()
With frmProducts
      If .cmbDepartments.Text = "<Unbound>" Then
        Depart = "0"
    Else
        Depart = Trim(Mid(.cmbDepartments.Text, 1, InStr(.cmbDepartments.Text, " -")))
    End If
    
    If .cmbScalePrefix.Text = "<None>" Then .cmbembedtype.Text = "<None>"
    If .cmbScalePrefix.Text <> "<None>" Then
    .chkTouch.Value = 0
    If .cmbembedtype.Text = "<None>" Then .cmbScalePrefix.Text = "<None>"
    End If
    If .cmbembedtype.Text = "Weight" Then .chkTouch.Value = 0
    If .cmbembedtype.Text = "Price" Then .chkTouch.Value = 0
    
    
    
    If .chkStock.Value = 0 Then .Txtstocklevel.Text = "None"
    If .chkScale.Value = 0 Then
    .cmbScalePrefix.Text = "<None>"
    .cmbembedtype.Text = "<None>"
    
    
    End If
    
    TaxRate = Mid(.cmbTax.Text, InStr(.cmbTax.Text, "-") + 2, InStr(.cmbTax.Text, "%") - InStr(.cmbTax.Text, "-") - 2)
    TaxType = Mid(.cmbTax.Text, 1, InStr(.cmbTax.Text, "-") - 2)
    ActiveReadServer "Select product_code from Products where Product_code='" & .txtProductCode.Text & "'"
    Select Case rs.RecordCount
        Case 0
            .grdProd.Rows = .grdProd.Rows + 1
            .grdProd.Tag = "1"
            .grdProd.Row = .grdProd.Rows - 1
            .grdProd.ShowCell .grdProd.Row, 0
            .grdProd.TextMatrix(.grdProd.Row, 0) = .txtProductCode.Text
            .grdProd.TextMatrix(.grdProd.Row, 1) = .txtDescription.Text & " " & Trim(.txtUnitSize) & .cmbUnit.Text
            .grdProd.TextMatrix(.grdProd.Row, 2) = .cmbDepartments.Text
            .grdProd.TextMatrix(.grdProd.Row, 3) = "0"
            .grdProd.TextMatrix(.grdProd.Row, 4) = Format(.txtLandCost.Text, "0.00")
            .grdProd.TextMatrix(.grdProd.Row, 5) = .cmbTax.Text
            .grdProd.TextMatrix(.grdProd.Row, 6) = Format(.txtSellIncl.Text, "0.00")
            .grdProd.Tag = ""
            
            ActiveUpdateServer "INSERT INTO [Products]([Product_Code], [Description], [Short_Description],[Nappi_Code]," & _
            "[Department_No], [Pack_Size], [Unit_Size], [Unit_of_Measure], [Maximum_Discount], [Sales_Item]," & _
            "[Stock_Item], [Returnable_Item], [Recipe_Item], [Touch_Item], [Scale_Item], [Landed_Cost]," & _
            "[Ave_Cost], [Selling_Price], [Sales_Tax], [Tax_Type], [Once_off], [Date_Created], [Date_Updated], [Scale_Prefix] ,[Kitchen1], [Kitchen2],[Weight_Full] ,[Weight_Empty], [Production_Item], [Scaleitemtype] )" & _
            "VALUES('" & .txtProductCode.Text & "','" & .txtDescription.Text & "','" & .txtShort.Text & "','" & .txtRef.Text & "','" & Depart & "'" & _
            ",'" & .txtPackSize.Text & "','" & .txtUnitSize.Text & "','" & .cmbUnit.Text & "',0," & .chkSales.Value & _
            "," & .chkStock.Value & "," & .chkDeposit.Value & "," & .chkRecipe.Value & "," & .chkTouch.Value & "," & .chkScale.Value & "," & Val(.txtLandCost.Text) & _
            "," & Val(.txtLandCost.Text) & "," & Val(.txtSellIncl.Text) & "," & TaxRate & "," & TaxType & "," & .chkDelete.Value & ",Getdate(),Getdate(),'" & .cmbScalePrefix.Text & "','" & .cmbPrinter1.Text & "','" & .cmbPrinter2.Text & "','" & Val(.txtFull.Text) & "','" & Val(.txtEmpty.Text) & "','" & .chkProduction.Value & "','" & .cmbembedtype.Text & "' )"
        Case 1
            If Val(Trim(Mid(.txtLandCost.ToolTipText, InStr(.txtLandCost.ToolTipText, ":") + 1))) = 0 Then
                AveCost = Val(.txtLandCost.Text)
                .txtLandCost.ToolTipText = " Average Cost: " & Format(Val(.txtLandCost.Text), "0.00") & " "
            Else
                ActiveReadServer1 "Select Sum(Stock_on_Hand) as Stock_on_Hand from Quantities where Product_Code = '" & .txtProductCode.Text & "'"
                If rs1.RecordCount > 0 Then
                    If Val(rs1.Fields("Stock_on_Hand") & "") = 0 Then
                        AveCost = .txtLandCost.Text
                    Else
                         AveCost = Val(Trim(Mid(.txtLandCost.ToolTipText, InStr(.txtLandCost.ToolTipText, ":") + 1)))
                    End If
                End If
                rs1.Close
            End If
            .grdProd.TextMatrix(.grdProd.Row, 0) = .txtProductCode.Text
            .grdProd.TextMatrix(.grdProd.Row, 1) = .txtDescription.Text & " " & Trim(.txtUnitSize) & .cmbUnit.Text
            .grdProd.TextMatrix(.grdProd.Row, 2) = .cmbDepartments.Text
            .grdProd.TextMatrix(.grdProd.Row, 3) = "0"
            .grdProd.TextMatrix(.grdProd.Row, 4) = Format(.txtLandCost.Text, "0.00")
            .grdProd.TextMatrix(.grdProd.Row, 5) = .cmbTax.Text
            .grdProd.TextMatrix(.grdProd.Row, 6) = Format(.txtSellIncl.Text, "0.00")
            ActiveUpdateServer "UPDATE [Products] SET " & _
            "[Description]='" & .txtDescription.Text & "'," & _
            "[Short_Description]='" & .txtShort.Text & "',[Nappi_Code]='" & .txtRef.Text & "'," & _
            "[Department_No]='" & Depart & "'," & _
            "[Pack_Size]='" & .txtPackSize.Text & "'," & _
            "[Unit_Size]='" & .txtUnitSize.Text & "'," & _
            "[Unit_of_Measure]='" & .cmbUnit.Text & "'," & _
            "[Maximum_Discount]=0," & _
            "[Sales_Item]=" & .chkSales.Value & "," & _
            "[Stock_Item]=" & .chkStock.Value & "," & _
            "[Returnable_Item]=" & .chkDeposit.Value & "," & _
            "[Recipe_Item]=" & .chkRecipe.Value & "," & _
            "[Touch_Item]=" & .chkTouch.Value & "," & _
            "[Production_Item]=" & .chkProduction.Value & "," & _
            "[Scale_Item]=" & .chkScale.Value & "," & _
            "[Whole_Unit]=" & .chkWhole.Value & "," & _
            "[Landed_Cost]=" & Val(.txtLandCost.Text) & "," & "[Ave_Cost]=" & AveCost & "," & _
            "[Selling_Price]=" & .txtSellIncl.Text & "," & _
            "[Sales_Tax]=" & TaxRate & "," & "[Tax_Type]=" & TaxType & "," & "[Stock_Level_Min]='" & Val(.Txtstocklevel.Text) & "'," & _
            "[Once_off]=" & .chkDelete.Value & ",[Date_Updated]=Getdate(),[Weight_Empty]='" & Val(.txtEmpty.Text) & "'," & _
            "[Scale_Prefix]='" & .cmbScalePrefix.Text & "'," & _
            "[Kitchen1]='" & .cmbPrinter1.Text & "',[Weight_Full]='" & Val(.txtFull.Text) & "'," & _
            "[Kitchen2] = '" & .cmbPrinter2.Text & "'," & "[Scaleitemtype] = '" & .cmbembedtype.Text & "'" & " WHERE Product_Code= '" & .txtProductCode.Text & "'"
            

    End Select
    rs.Close
    ActiveUpdateServer "Delete from Product_Prices where Product_Code='" & .txtProductCode.Text & "'"
    DoEvents
    ActiveUpdateServer "INSERT INTO [Product_Prices]([Product_Code], [Price2], [Price3], [Price4], [Price5],[Price6])" & _
    "VALUES ('" & .txtProductCode.Text & "'," & Val(.grdPrices.TextMatrix(7, 1)) & "," & Val(.grdPrices.TextMatrix(7, 2)) & "," & Val(.grdPrices.TextMatrix(7, 3)) & "," & Val(.grdPrices.TextMatrix(7, 4)) & "," & Val(.grdPrices.TextMatrix(7, 5)) & ")"
    DoEvents
    ActiveUpdateServer "Delete from Pack_Links where Product_Code='" & .txtProductCode.Text & "'"
    DoEvents
    If .txtLink.Text <> "<Not Linked>" Then
        ActiveUpdateServer "INSERT INTO [Pack_Links]([Product_Code], [Link_Code])" & _
        "VALUES ('" & .txtProductCode.Text & "','" & .txtLink.Text & "')"
    End If
    DoEvents
    ActiveUpdateServer "Delete from Recipes where Product_Code='" & .txtProductCode.Text & "'"
    DoEvents
    For i = 1 To .grdRecipe.Rows - 1
        If .grdRecipe.TextMatrix(i, 1) <> "" Then
            Select Case .grdRecipe.TextMatrix(i, 0)
                Case "Message"
                    LineType = 0
                Case "Preparation Recipe"
                    LineType = 1
                Case "Sales Item"
                    LineType = 2
                Case "Sales Item (Choice)"
                    LineType = 6
                Case "Stock Item"
                    LineType = 3
                Case "Stock Item (Choice)"
                    LineType = 7
                Case "Stock Item (Hidden)"
                    LineType = 4
                Case "Price/Size Change"
                    LineType = 5
                Case "Exit"
                    LineType = 8
                    .grdRecipe.TextMatrix(i, 1) = "Exit"
            End Select
            ActiveUpdateServer "INSERT INTO Recipes ([Line_Type], [Product_Code], [Line_Code], [Description], [Unit_of_Measure], [Qty_Used], [Cost])" & _
            "VALUES (" & LineType & ",'" & .txtProductCode.Text & "'," & Trim(Val(Trim(Mid(.grdRecipe.TextMatrix(i, 1), InStrRev(.grdRecipe.TextMatrix(i, 1), ",") + 1)))) & ",'" & .grdRecipe.TextMatrix(i, 1) & "','" & Trim(.grdRecipe.TextMatrix(i, 2)) & "','" & Trim(.grdRecipe.TextMatrix(i, 3)) & "'," & Val(.grdRecipe.TextMatrix(i, 4)) & ")"
        Else
            If .grdRecipe.TextMatrix(i, 0) = "Exit" Then
                LineType = 8
                .grdRecipe.TextMatrix(i, 1) = "Exit"
                ActiveUpdateServer "INSERT INTO Recipes ([Line_Type], [Product_Code], [Line_Code], [Description], [Unit_of_Measure], [Qty_Used], [Cost])" & _
                "VALUES (" & LineType & ",'" & .txtProductCode.Text & "'," & Trim(Val(Trim(Mid(.grdRecipe.TextMatrix(i, 1), InStrRev(.grdRecipe.TextMatrix(i, 1), ",") + 1)))) & ",'" & .grdRecipe.TextMatrix(i, 1) & "','" & Trim(.grdRecipe.TextMatrix(i, 2)) & "','" & Trim(.grdRecipe.TextMatrix(i, 3)) & "'," & Val(.grdRecipe.TextMatrix(i, 4)) & ")"
            End If
        End If
    Next i
    ActiveReadServer "Select * from Recipes where Line_Code= '" & .txtProductCode.Text & "' order by Line_No"
    While Not rs.EOF
        If rs.Fields("Description") <> .txtDescription.Text & " " & Trim(.txtUnitSize) & .cmbUnit.Text & " ," & .txtProductCode.Text Then
            ActiveUpdateServer "Update Recipes set Description=(select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size)" & _
            " + Unit_of_Measure END from Products where Products.Product_Code=Recipes.Line_Code)+' ,' + Line_Code from Recipes Where" & _
            " Line_Code = '" & .txtProductCode.Text & "'"
            DoEvents
            UpdateAveCost .txtProductCode.Text, Val(AveCost), 0, Val(.txtLandCost.Text)
        End If
        rs.MoveNext
    Wend
    rs.Close
    MsgBox "Product update successful", vbInformation, "HeroPOS"
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(4).Enabled = True
    frmMain.Toolbar1.Buttons(5).Enabled = True
End With
End Sub
Private Sub DeleteUsers()
With frmUsers
    Response = MsgBox("Are you sure you want to delete this user? " & Chr$(13) & "User No: " & .lblUserNumber.Caption & " - " & .txtFirstName.Text & " " & .txtFirstName.Text, vbYesNo, "HeroPOS")
        If Response = vbYes Then
            ActiveUpdateServer "Delete  from Users where User_no =" & .lblUserNumber.Caption
            .grdUsers.Rows = 1
            ActiveReadServer "Select User_no, User_Name, First_Name, Last_Name, isnull(Gender,0) as Gender, isnull(User_Type,0) as User_Type From Users order by User_No"
            While Not rs.EOF
                .grdUsers.Rows = .grdUsers.Rows + 1
                With .grdUsers
                    .TextMatrix(.Rows - 1, 0) = rs.Fields("User_No")
                    .TextMatrix(.Rows - 1, 1) = rs.Fields("User_Name") & ""
                    .TextMatrix(.Rows - 1, 2) = rs.Fields("First_Name") & " " & rs.Fields("Last_Name") & ""
                    Select Case rs.Fields("Gender")
                        Case False: .TextMatrix(.Rows - 1, 3) = "Male"
                        Case True: .TextMatrix(.Rows - 1, 3) = "Female"
                    End Select
                    Select Case rs.Fields("User_Type")
                        Case 0: .TextMatrix(.Rows - 1, 4) = "Manager"
                        Case 1: .TextMatrix(.Rows - 1, 4) = "Night Manager"
                        Case 2: .TextMatrix(.Rows - 1, 4) = "Reservations Clerk"
                        Case 3: .TextMatrix(.Rows - 1, 4) = "Waiter"
                        Case 4: .TextMatrix(.Rows - 1, 4) = "Barman"
                        Case 5: .TextMatrix(.Rows - 1, 4) = "GRV Clerk"
                        Case 6: .TextMatrix(.Rows - 1, 4) = "Buyer"
                        Case 7: .TextMatrix(.Rows - 1, 4) = "Supervisor"
                        Case 8: .TextMatrix(.Rows - 1, 4) = "Cashier"
                        Case 9: .TextMatrix(.Rows - 1, 4) = "Owner"
                        Case 10: .TextMatrix(.Rows - 1, 4) = "Staff Member"
                    End Select
                End With
                rs.MoveNext
            Wend
            rs.Close
            If .grdUsers.Rows > 1 Then .grdUsers.Row = 1
            frmMain.Toolbar1.Buttons(2).Enabled = True
            frmMain.Toolbar1.Buttons(4).Enabled = False
        End If
    End With
End Sub
Private Sub DeleteLocation()
With frmLocations
    Response = MsgBox("Are you sure you want to delete this Location? " & Chr$(13) & "Location No: " & frmLocations.txtLocNum.Text & " - " & frmLocations.txtLocName.Text, vbYesNo, "HeroPOS")
        If Response = vbYes Then
            ActiveUpdateServer "Delete  from Locations where Location_no =" & .txtLocNum.Text
            .grdLoc.Rows = 1
            ActiveReadServer "Select * From Locations order by Location_No"
            While Not rs.EOF
                .grdLoc.Rows = .grdLoc.Rows + 1
                .grdLoc.TextMatrix(.grdLoc.Rows - 1, 0) = rs.Fields("Location_No")
                .grdLoc.TextMatrix(.grdLoc.Rows - 1, 1) = rs.Fields("Loc_Name") & ""
                Select Case rs.Fields("Stock_Take")
                    Case 0: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 3) = "Enter Count"
                    Case 1: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 3) = "Use System Levels"
                    Case 2: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 3) = "Reset to Zero"
                End Select
                Select Case rs.Fields("Loc_Type")
                    Case 0: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 2) = "Sales Location"
                    Case 1: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 2) = "Stock Location"
                    Case 2: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 2) = "Expence Location"
                    Case 3: .grdLoc.TextMatrix(.grdLoc.Rows - 1, 2) = "Outside Location"
                End Select
                rs.MoveNext
            Wend
            rs.Close
            If .grdLoc.Rows > 1 Then .grdLoc.Row = 1
            frmMain.Toolbar1.Buttons(2).Enabled = True
            frmMain.Toolbar1.Buttons(4).Enabled = False
        End If
    End With
End Sub
Private Sub DeleteProduct()
On Error Resume Next
With frmProducts
    ActiveReadServer "Select * from Table_Listing where Product_Code = '" & frmProducts.txtProductCode.Text & "'"
    If rs.RecordCount > 0 Then
        MsgBox "You cannot delete this products as it is listed on an open table.", vbApplicationModal, "HeroPOS"
        rs.Close
        On Error Resume Next
        Exit Sub
    End If
    rs.Close
    ActiveReadServer "Select * from Tab_Listing where Product_Code = '" & frmProducts.txtProductCode.Text & "'"
    If rs.RecordCount > 0 Then
        MsgBox "You cannot delete this products as it is listed on an open bar tab.", vbApplicationModal, "HeroPOS"
        rs.Close
        On Error Resume Next
        Exit Sub
    End If
    rs.Close
    Response = MsgBox("Are you sure you want to delete this Product? " & Chr$(13) & "Product Code: " & frmProducts.txtProductCode.Text & " - " & frmProducts.txtDescription.Text, vbYesNo, "HeroPOS")
        If Response = vbYes Then
            ActiveReadServer "Select Product_Code,(Select Description from Products where Products.Product_Code=Recipes.Product_Code) as Description from Recipes where Line_Code = '" & frmProducts.txtProductCode.Text & "'"
            If rs.RecordCount > 0 Then
                Response = MsgBox("Deleting this Product will impact on " & rs.RecordCount & " other Products where it " & Chr$(13) & "is used in their Recipes. Are you sure you want to Continue?", vbYesNo, "HeroPOS")
                If Response = vbYes Then
                    MsgBox "A list of all the Products that were impacted by this action is stored as Delete.txt in the" & Chr$(13) & "Directory " & App.Path & "\Logs", vbCritical, "HeroPOS"
                    filenum = FreeFile
                    Open App.Path & "\Logs\Delete.txt" For Output As filenum
                    While Not rs.EOF
                        Print #filenum, rs.Fields("Product_Code") & " - " & rs.Fields("Description")
                        rs.MoveNext
                    Wend
                    Close filenum
                    ActiveUpdateServer "Delete from Recipes where Line_Code =" & .txtProductCode.Text
                Else
                    rs.Close
                    Exit Sub
                End If
            End If
            rs.Close
            ActiveUpdateServer "Delete from Recipes where Line_Code =" & .txtProductCode.Text
            DoEvents
            ActiveUpdateServer "Delete from Pack_Links where Product_Code ='" & .txtProductCode.Text & "'"
            DoEvents
            ActiveUpdateServer "Delete from Products where Product_Code ='" & .txtProductCode.Text & "'"
            .grdProd.RemoveItem .grdProd.Row
            ActiveReadServer "Select * from Products where Product_Code='" & .grdProd.TextMatrix(.grdProd.Row, 0) & "'"
            If rs.RecordCount > 0 Then
                .txtProductCode.Text = rs.Fields("Product_Code")
                .txtDescription.Text = rs.Fields("Description")
                .txtShort.Text = rs.Fields("Description")
                If rs.Fields("Unit_of_Measure") & "" = "" Then
                    .cmbUnit.Text = "each"
                Else
                    .cmbUnit.Text = rs.Fields("Unit_of_Measure")
                End If
                If rs.Fields("Unit_Size") = 0 Then
                    .txtUnitSize.Text = ""
                Else
                    .txtUnitSize.Text = rs.Fields("Unit_Size") & ""
                End If
                .cmbDepartments.Text = .grdProd.TextMatrix(.grdProd.Row, 2)
                .txtLandCost.Text = .grdProd.TextMatrix(.grdProd.Row, 4)
                .txtLandCost.ToolTipText = " Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & " "
                .cmbTax.Text = .grdProd.TextMatrix(.grdProd.Row, 5)
                .txtSellIncl.Text = .grdProd.TextMatrix(.grdProd.Row, 6)
                If rs.Fields("Scale_Prefix") & "" = "" Then
                    .cmbScalePrefix.Text = "20"
                Else
                    .cmbScalePrefix.Text = rs.Fields("Scale_Prefix") & ""
                End If
                If rs.Fields("Kitchen1") & "" = "" Then
                    .cmbPrinter1.Text = "<None>"
                Else
                    .cmbPrinter1.Text = rs.Fields("Kitchen1")
                End If
                If rs.Fields("Kitchen2") & "" = "" Then
                    .cmbPrinter2.Text = "<None>"
                Else
                    .cmbPrinter2.Text = rs.Fields("Kitchen2")
                End If
                .chkStock.Value = rs.Fields("Stock_Item")
                .chkSales.Value = rs.Fields("Sales_Item")
                .chkRecipe.Value = rs.Fields("Recipe_Item")
                .chkTouch.Value = rs.Fields("Touch_Item")
                .chkDeposit.Value = rs.Fields("Returnable_Item")
                .chkScale.Value = rs.Fields("Scale_Item")
                .chkDelete.Value = rs.Fields("Once_Off")
            End If
            rs.Close
            Select Case .chkRecipe.Value
                Case 0
                    .cmdTab(1).Enabled = False
                    .txtLandCost.Locked = False
                    .txtLandCost.Enabled = True
                    .lblCost.Caption = "Landed Cost:"
                Case 1
                    .cmdTab(1).Enabled = True
                    Select Case .cmbUnit.Text
                        Case "Preparation Recipe"
                            .txtLandCost.Enabled = False
                            .txtLandCost.Text = "N/A"
                        Case Else
                            .txtLandCost.Locked = True
                            .lblCost.Caption = "Theoretical Cost:"
                    End Select
                    ActiveReadServer "Select * from Recipes where Product_Code= '" & .txtProductCode.Text & "' order by Line_No"
                    While Not rs.EOF
                        .grdRecipe.Rows = .grdRecipe.Rows + 1
                        .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 1) = rs.Fields("Description")
                        Select Case rs.Fields("Line_Type")
                            Case 0
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 0) = "Message"
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 2) = " "
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 3) = " "
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 4) = " "
                                .grdRecipe.MergeRow(.grdRecipe.Rows - 1) = True
                                .grdRecipe.Cell(flexcpBackColor, .grdRecipe.Rows - 1, 2, .grdRecipe.Rows - 1, 4) = &HE0E0E0
                            Case 1
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 0) = "Preparation Recipe"
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 2) = " "
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 3) = " "
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 4) = " "
                                .grdRecipe.MergeRow(.grdRecipe.Rows - 1) = True
                                .grdRecipe.Cell(flexcpBackColor, .grdRecipe.Rows - 1, 2, .grdRecipe.Rows - 1, 4) = &HE0E0E0
                            Case 2
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 0) = "Sales Item"
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                                ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(.grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 1), InStrRev(.grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                                If rs1.RecordCount > 0 Then
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                                End If
                                rs1.Close
                            Case 3
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 0) = "Stock Item"
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                                ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(.grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 1), InStrRev(.grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                                If rs1.RecordCount > 0 Then
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                                End If
                                rs1.Close
                            Case 4
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 0) = "Stock Item (Hidden)"
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                                ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(.grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 1), InStrRev(.grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                                If rs1.RecordCount > 0 Then
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                                    .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                                End If
                                rs1.Close
                            Case 5
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 0) = "Price/Size Change"
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 2) = " "
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 3) = " "
                                .grdRecipe.TextMatrix(.grdRecipe.Rows - 1, 4) = " "
                                .grdRecipe.MergeRow(.grdRecipe.Rows - 1) = True
                                .grdRecipe.Cell(flexcpBackColor, .grdRecipe.Rows - 1, 2, .grdRecipe.Rows - 1, 4) = &HE0E0E0
                                rs1.Close
                        End Select
                        rs.MoveNext
                    Wend
                    rs.Close
            End Select
            If .grdRecipe.Rows = 1 Then .grdRecipe.Rows = 2
        End If
    End With
    On Error GoTo 0
End Sub
Private Sub DeleteDepartment()
On Error Resume Next
With frmDepartments
    Response = MsgBox("Are you sure you want to delete this Department? " & Chr$(13) & "Department No: " & frmDepartments.txtDep.Text & " - " & frmDepartments.txtDepName.Text, vbYesNo, "HeroPOS")
        If Response = vbYes Then
            ActiveUpdateServer "Delete  from Departments where Department_no ='" & .txtDep.Text & "'"
            Select Case .lblType.Caption
                Case "Major Department"
                    .grdMajor.Rows = 1
                    ActiveReadServer "Select * From Departments where Dept_Type=0 order by Department_No"
                    i = 0
                    While Not rs.EOF
                        .grdMajor.Rows = .grdMajor.Rows + 1
                        i = i + 1
                        .grdMajor.TextMatrix(i, 0) = rs.Fields("Department_No")
                        .grdMajor.TextMatrix(i, 1) = rs.Fields("Dept_Name")
                        rs.MoveNext
                    Wend
                    rs.Close
                    If .grdMajor.Rows > 0 Then
                        .grdMajor.Row = 1
                        .grdMajor.SetFocus
                    End If
                Case "Sub Department"
                    ActiveReadServer "Select * From Departments where Dept_Type=1 and Dept_Parent = '" & .grdMajor.TextMatrix(.grdMajor.Row, 0) & "' order by Department_No"
                    i = 0
                    .grdSub.Rows = 1
                    While Not rs.EOF
                        .grdSub.Rows = .grdSub.Rows + 1
                        i = i + 1
                        .grdSub.TextMatrix(i, 0) = rs.Fields("Department_No")
                        .grdSub.TextMatrix(i, 1) = rs.Fields("Dept_Name")
                        rs.MoveNext
                    Wend
                    rs.Close
                    If .grdSub.Rows > 0 Then
                        .grdSub.Row = 1
                        .grdSub.SetFocus
                    End If
            End Select
            frmMain.Toolbar1.Buttons(2).Enabled = True
            frmMain.Toolbar1.Buttons(4).Enabled = False
        End If
    End With
    On Error GoTo 0
End Sub

Public Sub CreateUser()
    If frmUsers.cmdUp.Caption = "6" Then
        frmUsers.Image2.top = 5310
        frmUsers.picTopbar.top = 5400
        frmUsers.grdUsers.top = 5940
        frmUsers.Picture1.top = 5430
        frmUsers.cmdUp.top = 5460
        frmUsers.cmdUp.Caption = "5"
        frmUsers.grdUsers.Height = 4260
    End If
    frmUsers.grdUsers.Rows = frmUsers.grdUsers.Rows + 1
    ActiveUpdateServer "INSERT INTO Users(User_Name, User_Password, First_Name, Last_Name, Expires, Gender, User_Type,  Ua_Reservations, Ua_Rooms, Ua_Guests, Ua_Checkin, Ua_Checkout, Ua_Users, Ua_Reports, Ua_Settings,Logged_In,All_Tables)" & _
    " VALUES('New User','','','',0,0,0,0,0,0,0,0,0,0,0,1,0)"
    ActiveReadServer "Select User_no,User_Name from Users where User_Name= 'New User'"
    If rs.RecordCount > 0 Then
        frmUsers.grdUsers.TextMatrix(frmUsers.grdUsers.Rows - 1, 0) = rs.Fields("User_No")
        frmUsers.grdUsers.TextMatrix(frmUsers.grdUsers.Rows - 1, 1) = rs.Fields("User_Name")
    End If
    rs.Close
    frmUsers.grdUsers.Row = frmUsers.grdUsers.Rows - 1
    frmUsers.txtLogin.SetFocus
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    frmUsers.btnSetting.Enabled = False
End Sub

Public Sub SaveDetails()
With frmDetails
    Select Case .cmbType.Text
        Case "Guest House": Vtype = 1
        Case "Boutique Hotel": Vtype = 2
        Case "Lodge": Vtype = 3
        Case "Hotel": Vtype = 4
        Case "Restaurant": Vtype = 5
        Case "Guest House and Restaurant": Vtype = 6
        Case "Club": Vtype = 7
        Case "Supermarket": Vtype = 8
        Case "Butchery": Vtype = 9
        Case "Warehouse": Vtype = 10
        Case "Filling Station": Vtype = 11
        Case "Retail Store": Vtype = 12
    End Select
   
    ActiveUpdateServer "Update Branch_Details set " & _
            "Branch_No = " & .txtNo.Text & _
            ",Branch_Name='" & .txtName.Text & "'" & _
            ",Address='" & .txtAddress.Text & "'" & _
            ",Region=" & Val(Mid(.cmbRegion.Text, 1, InStr(.cmbRegion.Text, "-") - 1)) & _
            ",Branch_Type=" & Vtype & _
            ",Prefix='" & .txtPrefix.Text & "'" & _
            ",Tax_Rate='" & .lblTax.Caption & "'" & _
            ",Time_Start='" & .tmPicker(0).Value & "'" & _
            ",Time_Stop='" & .tmPicker(1).Value & "'" & _
            ",Fin_Year='" & .dtPicker.Value & "'" & _
            " where Branch_no =" & .txtNo.Text
        MsgBox "Branch Details update successful", vbInformation
End With
End Sub
Public Sub DeleteDebtor()
    If Val(frmGuests.grdSupp.TextMatrix(frmGuests.grdSupp.Row, 4)) <> 0 Then
        MsgBox "You cannot delete a debtor with and outstanding balance." & Chr(13) & "Clear the outstanding balance before deleting", vbInformation, "HeroPOS"
    Else
        Response = MsgBox("Are you certain you want to delete this debtor." & Chr(13) & "Deleting a debtor clears all history relating to the debtor.", vbYesNo, "HeroPOS")
        Select Case Response
            Case vbYes
                ActiveUpdateServer "Delete from Debtors where Debtor_No = '" & frmGuests.grdSupp.TextMatrix(frmGuests.grdSupp.Row, 0) & "'"
                frmGuests.grdSupp.RemoveItem frmGuests.grdSupp.Row
            Case vbNo
        End Select
    End If
End Sub
Public Sub DeleteRoom()
    ActiveReadServer "Select * from Room_Accounts where Account_No = " & frmRooms.grdRooms.TextMatrix(frmRooms.grdRooms.Row, 0)
    If rs.RecordCount > 0 Then
        If Round(rs.Fields("Balance"), 2) <> 0 Then
            MsgBox "You cannot delete a Room with and outstanding balance." & Chr(13) & "Clear the outstanding balance before deleting", vbInformation, "HeroPOS"
            Exit Sub
        End If
    Else
        Response = MsgBox("Are you certain you want to delete this room.", vbYesNo, "HeroPOS")
        Select Case Response
            Case vbYes
                ActiveUpdateServer "Delete from Rooms where Room_No = '" & frmRooms.grdRooms.TextMatrix(frmRooms.grdRooms.Row, 0) & "'"
                frmRooms.grdRooms.RemoveItem frmRooms.grdRooms.Row
            Case vbNo
        End Select
    End If
End Sub
Public Sub DeleteSupplier()
    If Val(frmSuppliers.grdSupp.TextMatrix(frmSuppliers.grdSupp.Row, 4)) <> 0 Then
        MsgBox "You cannot delete a supplier with and outstanding balance." & Chr(13) & "Clear the outstanding balance before deleting", vbInformation, "HeroPOS"
    Else
        Response = MsgBox("Are you certain you want to delete this supplier." & Chr(13) & "Deleting a supplier clears all history relating to the supplier.", vbYesNo, "HeroPOS")
        Select Case Response
            Case vbYes
                ActiveUpdateServer "Delete from Suppliers where Supplier_No = '" & frmSuppliers.grdSupp.TextMatrix(frmSuppliers.grdSupp.Row, 0) & "'"
                frmSuppliers.grdSupp.RemoveItem frmSuppliers.grdSupp.Row
            Case vbNo
        End Select
    End If
End Sub
Private Sub Trade_Print_Slip()
With frmReports
    On Error GoTo trap
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, .mthViewEnd.Value)
        .lblDate.Caption = Format(.mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = .mthViewEnd.Value
    End If
    PrintErr = 0
    Slip_Port = ""
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then  'Kotie 17-03-20-13
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
                If Slip_Port = "FILE:" Then
                    Open "C:\" & x.DeviceName & ".txt" For Output As filenum
                Else
                    Open Slip_Port For Output As filenum
                End If
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(1);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, UCase(Branch_Name)
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "TRADE ANALYSIS"
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, frmReports.lblDate.Caption
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
    
    Print #filenum, String(40, "=")
    Print #filenum, "Spend per Head: " & frmReports.grdGP.TextMatrix(1, 3) & String(12 - Len(Str(frmReports.grdGP.TextMatrix(1, 3))), " ")
    Print #filenum, "Customer Count: " & frmReports.grdGP.TextMatrix(0, 3) & String(12 - Len(Str(frmReports.grdGP.TextMatrix(0, 3))), " ")
    Print #filenum, String(40, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, "Cash Sales: " & String(9 - Len(.grdRev.TextMatrix(1, 1)), " ") & .grdRev.TextMatrix(1, 1) & "   " & .grdRev.TextMatrix(1, 2) & String(4 - Len(.grdRev.TextMatrix(1, 2)), " ")
    Print #filenum, "Card Sales: " & String(9 - Len(.grdRev.TextMatrix(2, 1)), " ") & .grdRev.TextMatrix(2, 1) & "   " & .grdRev.TextMatrix(2, 2) & String(4 - Len(.grdRev.TextMatrix(2, 2)), " ")
    Print #filenum, "Voucher Sales: " & String(9 - Len(.grdRev.TextMatrix(3, 1)), " ") & .grdRev.TextMatrix(3, 1) & "   " & .grdRev.TextMatrix(3, 2) & String(4 - Len(.grdRev.TextMatrix(3, 2)), " ")
    Print #filenum, "Charge Sales: " & String(9 - Len(.grdRev.TextMatrix(4, 1)), " ") & .grdRev.TextMatrix(4, 1) & "   " & .grdRev.TextMatrix(4, 2) & String(4 - Len(.grdRev.TextMatrix(4, 2)), " ")
    Print #filenum, "Loyalty Sales: " & String(9 - Len(.grdRev.TextMatrix(5, 1)), " ") & .grdRev.TextMatrix(5, 1) & "   " & .grdRev.TextMatrix(5, 2) & String(4 - Len(.grdRev.TextMatrix(5, 2)), " ")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(51) & Chr(18);
    Print #filenum, ""
    Print #filenum, "TOTAL: " & String(9 - Len(.grdRev.TextMatrix(6, 1)), " ") & .grdRev.TextMatrix(6, 1) & "   " & .grdRev.TextMatrix(6, 2) & String(4 - Len(.grdRev.TextMatrix(6, 2)), " ")
    Print #filenum, ""
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(50);
    Print #filenum, "-Payouts: " & String(9 - Len(.grdTrans.TextMatrix(1, 1)), " ") & .grdTrans.TextMatrix(1, 1) & "   " & .grdTrans.TextMatrix(1, 2) & String(4 - Len(.grdTrans.TextMatrix(1, 2)), " ")
    Print #filenum, "+Deposits: " & String(9 - Len(.grdTrans.TextMatrix(2, 1)), " ") & .grdTrans.TextMatrix(2, 1) & "   " & .grdTrans.TextMatrix(2, 2) & String(4 - Len(.grdTrans.TextMatrix(2, 2)), " ")
    Print #filenum, "+Receive on Acc: " & String(9 - Len(.grdTrans.TextMatrix(3, 1)), " ") & .grdTrans.TextMatrix(3, 1) & "   " & .grdTrans.TextMatrix(3, 2) & String(4 - Len(.grdTrans.TextMatrix(3, 2)), " ")
    Print #filenum, String(33, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "SUB TOTAL: " & String(9 - Len(.grdTrans.TextMatrix(4, 1)), " ") & .grdTrans.TextMatrix(4, 1)
    Print #filenum, "REPORTED: " & String(9 - Len(.grdTrans.TextMatrix(5, 1)), " ") & .grdTrans.TextMatrix(5, 1)
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, String(33, "=")
    Print #filenum, "Item corrects: " & String(9 - Len(.grdCount.TextMatrix(1, 1)), " ") & .grdCount.TextMatrix(1, 1) & "   " & .grdCount.TextMatrix(1, 2) & String(4 - Len(.grdCount.TextMatrix(1, 2)), " ")
    Print #filenum, "Voids: " & String(9 - Len(.grdCount.TextMatrix(2, 1)), " ") & .grdCount.TextMatrix(2, 1) & "   " & .grdCount.TextMatrix(2, 2) & String(4 - Len(.grdCount.TextMatrix(2, 2)), " ")
    Print #filenum, "Returns: " & String(9 - Len(.grdCount.TextMatrix(3, 1)), " ") & .grdCount.TextMatrix(3, 1) & "   " & .grdCount.TextMatrix(3, 2) & String(4 - Len(.grdCount.TextMatrix(3, 2)), " ")
    Print #filenum, "Wastages: " & String(9 - Len(.grdCount.TextMatrix(4, 1)), " ") & .grdCount.TextMatrix(4, 1) & "   " & .grdCount.TextMatrix(4, 2) & String(4 - Len(.grdCount.TextMatrix(4, 2)), " ")
    Print #filenum, "Discount%: " & String(9 - Len(.grdCount.TextMatrix(5, 1)), " ") & .grdCount.TextMatrix(5, 1) & "   " & .grdCount.TextMatrix(5, 2) & String(4 - Len(.grdCount.TextMatrix(5, 2)), " ")
    Print #filenum, "Discount Value: " & String(9 - Len(.grdCount.TextMatrix(6, 1)), " ") & .grdCount.TextMatrix(6, 1) & "   " & .grdCount.TextMatrix(6, 2) & String(4 - Len(.grdCount.TextMatrix(6, 2)), " ")
    Print #filenum, "Service Charges: " & String(9 - Len(.grdCount.TextMatrix(7, 1)), " ") & .grdCount.TextMatrix(7, 1) & "   " & .grdCount.TextMatrix(7, 2) & String(4 - Len(.grdCount.TextMatrix(7, 2)), " ")
    Print #filenum, String(33, "=")
    Print #filenum, String(9, "=")
    Print #filenum, String(33, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "DEPARTMENT BREAKDOWN"
    Print #filenum, "--------------------"
    ActiveReadServer3 "SELECT LEFT(Sales_Journal.Department_No, PATINDEX('%-%', Sales_Journal.Department_No) - 1) AS Department_No," & _
    " (Select Dept_Name from Departments where Department_No =LEFT(Sales_Journal.Department_No, PATINDEX('%-%', Sales_Journal.Department_No) - 1))" & _
    " as Department_Name, SUM(Line_Total) AS Line_Total" & _
    " From dbo.Sales_Journal" & _
    " WHERE (isnull(Department_No,'')<>'') and (Line_Total <> 0 ) and (Date_Time > '" & .mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') AND (Function_Key = 7) AND (Extra <> 'Corr')" & _
    " GROUP BY LEFT(Department_No, PATINDEX('%-%', Department_No) - 1)"
    While Not rs3.EOF
        Print #filenum, rs3.Fields("Department_Name") & ": " & String(9 - Len(Format(rs3.Fields("Line_Total"), "0.00")), " ") & Format(rs3.Fields("Line_Total"), "0.00")
        rs3.MoveNext
    Wend
    rs3.Close
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(1);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, "PRESENTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "ACCEPTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "DATED:"
    Print #filenum, ""
    Print #filenum, Chr(27) & Chr(50);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(100) & Chr(7);
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(1);
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #1
    On Error GoTo 0
    frmTillReport.Tag = ""
On Error GoTo 0
End With
Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
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


Private Sub EOD_Slip()

With frmReports
    On Error GoTo trap
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, .mthViewEnd.Value)
        .lblDate.Caption = Format(.mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = .mthViewEnd.Value
    End If
    PrintErr = 0
    Slip_Port = ""
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then  'Kotie 17-03-20-13
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
                If Slip_Port = "FILE:" Then
                    Open "C:\" & x.DeviceName & ".txt" For Output As filenum
                Else
                    Open Slip_Port For Output As filenum
                End If
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(1);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, UCase(Branch_Name)
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "TRADE ANALYSIS"
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(33) & Chr(0); 'Normal Font
    Print #filenum, frmReports.lblDate.Caption
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
    
    
    
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49); 'Dark Font
    
    
    With frmReports
    ' Revenue
        Total = 0
        Cnt = 0
        Head = "Revenue"
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49); 'Dark Font
        Print #filenum, Head & String(42 - Len(Head), " ")
        Print #filenum, Chr(27) & Chr(33) & Chr(0); 'Normal Font
        For i = 1 To .grdRev.Rows - 1
            strHeading = .grdRev.TextMatrix(i, 0)
            strVal = Format(.grdRev.TextMatrix(i, 1), "0.00")
            strcnt = .grdRev.TextMatrix(i, 2)
            Print #filenum, strHeading & String(25 - (Len(strHeading)), " ") & _
                            String(10 - (Len(strVal)), " ") & strVal & _
                             String(6 - (Len(strcnt)), " ") & strcnt
            Total = Total + .grdRev.ValueMatrix(i, 1)
            Total = Cnt + .grdRev.ValueMatrix(i, 2)
        Next i
        Print #filenum,
        'Print #filenum, strHeading & String(20 - (Len("Revenue Total:")), " ") & _
        '                strVal & String(10 - (Len(Total)), " ") & _
                        strcnt & String(6 - (Len(Total)), " ")
                        
    'Transactions
        Total = 0
        Cnt = 0
        Head = "Revenue Transactions"
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49); 'Dark Font
        Print #filenum, Head & String(42 - Len(Head), " ")
        Print #filenum, Chr(27) & Chr(33) & Chr(0); 'Normal Font
        For i = 1 To .grdTrans.Rows - 1
            strHeading = .grdTrans.TextMatrix(i, 0)
            strVal = Format(.grdTrans.TextMatrix(i, 1), "0.00")
            strcnt = .grdTrans.TextMatrix(i, 2)
            Print #filenum, strHeading & String(25 - (Len(strHeading)), " ") & _
                            String(10 - (Len(strVal)), " ") & strVal & _
                            String(6 - (Len(strcnt)), " ") & strcnt
            Total = Total + .grdRev.ValueMatrix(i, 1)
            Total = Cnt + .grdRev.ValueMatrix(i, 2)
        Next i
        Print #filenum,
       ' Print #filenum, strHeading & String(26 - (Len("Total Reported:")), " ") & _
       '                 strVal & String(10 - (Len(Total)), " ") & _
       '                 strcnt & String(6 - (Len(Total)), " ")
                        
    ' Other Transactions
        Total = 0
        Cnt = 0
        Head = "Other Transactions"
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49); 'Dark Font
        Print #filenum, Head & String(42 - Len(Head), " ")
        Print #filenum, Chr(27) & Chr(33) & Chr(0); 'Normal Font
        For i = 1 To .grdCount.Rows - 1
            strHeading = .grdCount.TextMatrix(i, 0)
            strVal = Format(.grdCount.TextMatrix(i, 1), "0.00")
            strcnt = .grdCount.TextMatrix(i, 2)
            Print #filenum, strHeading & String(25 - (Len(strHeading)), " ") & _
                            String(10 - (Len(strVal)), " ") & strVal & _
                            String(6 - (Len(strcnt)), " ") & strcnt
        Next i
        Print #filenum,
                        
        
    ' Tax cnters
        Total = 0
        Head = "Tax Counters"
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49); 'Dark Font
        Print #filenum, Head & String(42 - Len(Head), " ")
        Print #filenum, Chr(27) & Chr(33) & Chr(0); 'Normal Font
        For i = 1 To .grdTax.Rows - 1
            strHeading = .grdTax.TextMatrix(i, 0)
            strVal = Format(.grdTax.TextMatrix(i, 1), "0.00")
            Print #filenum, strHeading & String(25 - (Len(strHeading)), " ") & _
                             String(10 - (Len(strVal)), " ") & strVal & String(6, " ")
        Next i
        Print #filenum,
        
    ' Stock
        Total = 0
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49); 'Dark Font
        Head = "Stock"
        Print #filenum, Head & String(42 - Len(Head), " ")
        Print #filenum, Chr(27) & Chr(33) & Chr(0); 'Normal Font
        For i = 1 To .grdStock.Rows - 1
            strHeading = .grdStock.TextMatrix(i, 0)
            strVal = Format(.grdStock.TextMatrix(i, 1), "0.00")
            Print #filenum, strHeading & String(25 - (Len(strHeading)), " ") & _
                            String(10 - (Len(strVal)), " ") & strVal & String(6, " ")
        Next i
        Print #filenum,
        
    ' Debtors
        Total = 0
        Head = "Debtors"
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49); 'Dark Font
        Print #filenum, Head & String(42 - Len(Head), " ")
        Print #filenum, Chr(27) & Chr(33) & Chr(0); 'Normal Font
        For i = 1 To .grdCred.Rows - 1
            strHeading = .grdCred.TextMatrix(i, 0)
            strVal = Format(.grdCred.TextMatrix(i, 1), "0.00")
            Print #filenum, strHeading & String(25 - (Len(strHeading)), " ") & _
                            String(10 - (Len(strVal)), " ") & strVal & String(6, " ")
        Next i
        Print #filenum,
                        
        
    ' Ceditors
        Total = 0
        Head = "Creditors"
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49); 'Dark Font
        Print #filenum, Head & String(42 - Len(Head), " ")
        Print #filenum, Chr(27) & Chr(33) & Chr(0); 'Normal Font
        For i = 1 To .grdCred.Rows - 1
            strHeading = .grdCred.TextMatrix(i, 2)
            strVal = Format(.grdCred.TextMatrix(i, 3), "0.00")
            Print #filenum, strHeading & String(25 - (Len(strHeading)), " ") & _
                            String(10 - (Len(strVal)), " ") & strVal & String(6, " ")
        Next i
        Print #filenum,
                        
                        
    End With
    
    Print #filenum, "Gross Frofit %:" & String(16, " ") & frmReports.grdGP.TextMatrix(0, 1) & String(11 - Len(frmReports.grdGP.TextMatrix(0, 1)), " ")
    Print #filenum, "Gross Profit Value:" & String(11, " ") & frmReports.grdGP.TextMatrix(1, 1) & String(13 - Len(Str(frmReports.grdGP.TextMatrix(1, 1))), " ")
    Print #filenum, "Customer Count:" & String(20, " ") & frmReports.grdGP.TextMatrix(0, 3) & String(8 - Len(Str(frmReports.grdGP.TextMatrix(0, 3))), " ")
    Print #filenum, "Spend per Head:" & String(17, " ") & frmReports.grdGP.TextMatrix(1, 3) & String(8 - Len(Str(frmReports.grdGP.TextMatrix(1, 3))), " ")

    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum,
    Print #filenum, "DEPARTMENT BREAKDOWN           "
    Print #filenum, String(42, "_")
    ActiveReadServer3 "SELECT LEFT(Sales_Journal.Department_No, PATINDEX('%-%', Sales_Journal.Department_No) - 1) AS Department_No," & _
    " (Select Dept_Name from Departments where Department_No =LEFT(Sales_Journal.Department_No, PATINDEX('%-%', Sales_Journal.Department_No) - 1))" & _
    " as Department_Name, SUM(Line_Total) AS Line_Total" & _
    " From dbo.Sales_Journal" & _
    " WHERE (isnull(Department_No,'')<>'') and (Line_Total <> 0 ) and (Date_Time > '" & .mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') AND (Function_Key = 7) AND (Extra <> 'Corr')" & _
    " GROUP BY LEFT(Department_No, PATINDEX('%-%', Department_No) - 1)"
    While Not rs3.EOF
        Print #filenum, rs3.Fields("Department_Name") & ":        " & _
        String(9 - Len(Format(rs3.Fields("Line_Total"), "0.00")), " ") & _
        Format(rs3.Fields("Line_Total"), "0.00") & "      "
        rs3.MoveNext
    Wend
    Print #filenum, String(42, "_")
    ActiveReadServer3 ("select Dept_name, sum(Line_total * Conversion_rate) as Line_total from Sales_by_department_forreign_currency " & _
    " where (Date_Time > '" & .mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "'" & _
    " and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " group by Dept_name")
    Print #filenum,
    Print #filenum,
    
    Cnt = 0
    While Not rs3.EOF
        If Cnt = 0 Then
            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
            Print #filenum, "Foreign Currency             "
            Print #filenum, String(42, "_")
            Cnt = 1
        End If
        Print #filenum, rs3.Fields("Dept_name") & ":        " & _
        String(9 - Len(Format(rs3.Fields("Line_Total"), "0.00")), " ") & _
        Format(rs3.Fields("Line_Total"), "0.00") & "      "
        rs3.MoveNext
    Wend
    Print #filenum,
    Print #filenum,
    Print #filenum,
    Print #filenum,
    Print #filenum,
    Print #filenum, Chr(29) & "V" & Chr(49);

Close #filenum
End With
Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
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

