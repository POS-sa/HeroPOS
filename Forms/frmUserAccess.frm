VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmUserAccess 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " User Access Rights"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   Icon            =   "frmUserAccess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picPage2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   30
      ScaleHeight     =   5055
      ScaleWidth      =   8835
      TabIndex        =   9
      Top             =   1110
      Visible         =   0   'False
      Width           =   8835
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   29
         Left            =   6240
         TabIndex        =   73
         Top             =   3480
         Width           =   2295
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4048;609"
         Value           =   "0"
         Caption         =   "Tab/table owner transfer"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   28
         Left            =   6240
         TabIndex        =   72
         Top             =   1605
         Width           =   2445
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4313;609"
         Value           =   "0"
         Caption         =   "Bill Reprint"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   27
         Left            =   6240
         TabIndex        =   71
         Top             =   3000
         Width           =   2295
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4048;609"
         Value           =   "0"
         Caption         =   "Quotes"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   26
         Left            =   6240
         TabIndex        =   70
         Top             =   2535
         Width           =   2445
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4313;609"
         Value           =   "0"
         Caption         =   "Service Charges"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   25
         Left            =   6240
         TabIndex        =   64
         Top             =   240
         Width           =   2445
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4313;609"
         Value           =   "0"
         Caption         =   "No Sales"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   24
         Left            =   6240
         TabIndex        =   60
         Top             =   690
         Width           =   2445
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4313;609"
         Value           =   "0"
         Caption         =   "Wastages"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   23
         Left            =   6240
         TabIndex        =   53
         Top             =   3945
         Width           =   2445
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4313;609"
         Value           =   "0"
         Caption         =   "Application Exit"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   22
         Left            =   6240
         TabIndex        =   52
         Top             =   4410
         Width           =   2445
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4313;609"
         Value           =   "0"
         Caption         =   "All System Total Clear"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   21
         Left            =   6240
         TabIndex        =   41
         Top             =   2070
         Width           =   2445
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4313;609"
         Value           =   "0"
         Caption         =   "Product Searches"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   20
         Left            =   6240
         TabIndex        =   40
         Top             =   1155
         Width           =   2445
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4313;609"
         Value           =   "0"
         Caption         =   "Buffer Prints"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   19
         Left            =   3315
         TabIndex        =   39
         Top             =   4395
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Price Overrides"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   18
         Left            =   3315
         TabIndex        =   38
         Top             =   3930
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Table and Tab Transfers"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   17
         Left            =   3315
         TabIndex        =   37
         Top             =   3465
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Transaction Clears"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   16
         Left            =   3315
         TabIndex        =   36
         Top             =   3000
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Split Tendering"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   15
         Left            =   3315
         TabIndex        =   35
         Top             =   2535
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Sales Store and Retreive"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   14
         Left            =   3315
         TabIndex        =   34
         Top             =   2070
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Receive on Accounts"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   13
         Left            =   3315
         TabIndex        =   33
         Top             =   1605
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Loans"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   12
         Left            =   3315
         TabIndex        =   32
         Top             =   1140
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Cash Pick Ups"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   11
         Left            =   3315
         TabIndex        =   31
         Top             =   675
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Payouts"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   10
         Left            =   3315
         TabIndex        =   30
         Top             =   210
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   " Over Tenders"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   9
         Left            =   480
         TabIndex        =   29
         Top             =   4395
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Discount Amount"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   8
         Left            =   480
         TabIndex        =   28
         Top             =   3930
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Discount Percentage"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   7
         Left            =   480
         TabIndex        =   27
         Top             =   3465
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Returns"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   6
         Left            =   480
         TabIndex        =   26
         Top             =   3000
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Voids"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   5
         Left            =   480
         TabIndex        =   25
         Top             =   2535
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Item Corrects"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   4
         Left            =   480
         TabIndex        =   24
         Top             =   2070
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Loyalty Sales"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   3
         Left            =   480
         TabIndex        =   23
         Top             =   1605
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Charge Sales"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   2
         Left            =   480
         TabIndex        =   22
         Top             =   1140
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Card Sales"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   675
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4410;609"
         Value           =   "0"
         Caption         =   "Voucher Sales"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkBox 
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   210
         Width           =   2505
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4419;609"
         Value           =   "0"
         Caption         =   "Cash Sales"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image Image2 
         Height          =   4995
         Left            =   120
         Top             =   0
         Width           =   8595
         BorderColor     =   12632256
         BackColor       =   16777215
         Size            =   "15161;8811"
         VariousPropertyBits=   19
      End
   End
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   6240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   1
      Left            =   7410
      TabIndex        =   2
      Top             =   6240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Help"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   2
      Left            =   4590
      TabIndex        =   0
      Top             =   6240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   420
      Index           =   1
      Left            =   2175
      TabIndex        =   6
      Top             =   720
      Width           =   2085
      _Version        =   524298
      _ExtentX        =   3678
      _ExtentY        =   741
      _StockProps     =   66
      Caption         =   "Application Rights"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      Shape           =   4
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmUserAccess.frx":000C
      textLT          =   "frmUserAccess.frx":0090
      textCT          =   "frmUserAccess.frx":00A8
      textRT          =   "frmUserAccess.frx":00C0
      textLM          =   "frmUserAccess.frx":00D8
      textRM          =   "frmUserAccess.frx":00F0
      textLB          =   "frmUserAccess.frx":0108
      textCB          =   "frmUserAccess.frx":0120
      textRB          =   "frmUserAccess.frx":0138
      colorBack       =   "frmUserAccess.frx":0150
      colorIntern     =   "frmUserAccess.frx":017A
      colorMO         =   "frmUserAccess.frx":01A4
      colorFocus      =   "frmUserAccess.frx":01CE
      colorDisabled   =   "frmUserAccess.frx":01F8
      colorPressed    =   "frmUserAccess.frx":0222
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   5
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   420
      Index           =   0
      Left            =   150
      TabIndex        =   7
      Top             =   720
      Width           =   2085
      _Version        =   524298
      _ExtentX        =   3678
      _ExtentY        =   741
      _StockProps     =   66
      Caption         =   "Application Access"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
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
      Shape           =   4
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmUserAccess.frx":024C
      textLT          =   "frmUserAccess.frx":02D0
      textCT          =   "frmUserAccess.frx":02E8
      textRT          =   "frmUserAccess.frx":0300
      textLM          =   "frmUserAccess.frx":0318
      textRM          =   "frmUserAccess.frx":0330
      textLB          =   "frmUserAccess.frx":0348
      textCB          =   "frmUserAccess.frx":0360
      textRB          =   "frmUserAccess.frx":0378
      colorBack       =   "frmUserAccess.frx":0390
      colorIntern     =   "frmUserAccess.frx":03BA
      colorMO         =   "frmUserAccess.frx":03E4
      colorFocus      =   "frmUserAccess.frx":040E
      colorDisabled   =   "frmUserAccess.frx":0438
      colorPressed    =   "frmUserAccess.frx":0462
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   5
      Value           =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   420
      Index           =   2
      Left            =   4230
      TabIndex        =   42
      Top             =   720
      Width           =   2085
      _Version        =   524298
      _ExtentX        =   3678
      _ExtentY        =   741
      _StockProps     =   66
      Caption         =   "Settings"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      Shape           =   4
      Surface         =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmUserAccess.frx":048C
      textLT          =   "frmUserAccess.frx":04FC
      textCT          =   "frmUserAccess.frx":0514
      textRT          =   "frmUserAccess.frx":052C
      textLM          =   "frmUserAccess.frx":0544
      textRM          =   "frmUserAccess.frx":055C
      textLB          =   "frmUserAccess.frx":0574
      textCB          =   "frmUserAccess.frx":058C
      textRB          =   "frmUserAccess.frx":05A4
      colorBack       =   "frmUserAccess.frx":05BC
      colorIntern     =   "frmUserAccess.frx":05E6
      colorMO         =   "frmUserAccess.frx":0610
      colorFocus      =   "frmUserAccess.frx":063A
      colorDisabled   =   "frmUserAccess.frx":0664
      colorPressed    =   "frmUserAccess.frx":068E
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   5
   End
   Begin VB.PictureBox picPage3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   150
      ScaleHeight     =   5025
      ScaleWidth      =   8625
      TabIndex        =   43
      Top             =   1110
      Visible         =   0   'False
      Width           =   8625
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Commision Calculation..."
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   4350
         TabIndex        =   61
         Top             =   90
         Width           =   4125
         Begin VB.Label lblSub 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Departments (Commisions)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   30
            TabIndex        =   67
            Top             =   1350
            Width           =   4035
         End
         Begin MSForms.ComboBox cmbDepartments 
            Height          =   285
            Left            =   1410
            TabIndex        =   66
            Tag             =   "Up"
            Top             =   870
            Width           =   2625
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "4630;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label2 
            Height          =   225
            Index           =   3
            Left            =   30
            TabIndex        =   65
            Top             =   930
            Width           =   1365
            BackColor       =   -2147483643
            Caption         =   "Major Department:"
            Size            =   "2408;397"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Line Line1 
            X1              =   4110
            X2              =   0
            Y1              =   780
            Y2              =   780
         End
         Begin MSForms.CheckBox chkComms 
            Height          =   375
            Left            =   150
            TabIndex        =   62
            Tag             =   "1"
            Top             =   300
            Width           =   2505
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "4419;661"
            Value           =   "0"
            Caption         =   "On Turnover Excluding Tax"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Image picMin 
            Height          =   495
            Left            =   0
            Top             =   1230
            Width           =   4125
            Size            =   "7276;873"
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Drawer Open for..."
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   240
         TabIndex        =   54
         Top             =   2520
         Width           =   2265
         Begin MSForms.CheckBox chkPanels 
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   59
            Tag             =   "1"
            Top             =   1830
            Width           =   1905
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;503"
            Value           =   "0"
            Caption         =   "Loyalty Sales"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPanels 
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   58
            Tag             =   "1"
            Top             =   1410
            Width           =   1905
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;503"
            Value           =   "0"
            Caption         =   "Charge Sales"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPanels 
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   57
            Tag             =   "1"
            Top             =   1050
            Width           =   1905
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;503"
            Value           =   "0"
            Caption         =   "Voucher Sales"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPanels 
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   56
            Tag             =   "1"
            Top             =   675
            Width           =   1905
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;503"
            Value           =   "0"
            Caption         =   "Card Sales"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPanels 
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   55
            Tag             =   "1"
            Top             =   300
            Width           =   1905
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;503"
            Value           =   "0"
            Caption         =   "Cash Sales"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid grdSub 
         Height          =   3105
         Left            =   4350
         TabIndex        =   68
         Top             =   1800
         Width           =   4120
         _cx             =   7267
         _cy             =   5477
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
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
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   700
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmUserAccess.frx":06B8
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
      Begin MSForms.CheckBox chkBarCash 
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   2100
         Width           =   3975
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "7011;661"
         Value           =   "0"
         Caption         =   "User can Finalize all Open Tables and Tabs"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkAllTables 
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   1800
         Width           =   3975
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "7011;661"
         Value           =   "0"
         Caption         =   "User can see all open Tables or Tabs from all Users"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkLog 
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   1470
         Width           =   2685
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4736;661"
         Value           =   "0"
         Caption         =   "User Stays Logged in "
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   49
         Top             =   1080
         Width           =   2115
         BackColor       =   16777215
         Caption         =   "Deduction% from Card Tipps:"
         Size            =   "3731;450"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtComm2 
         Height          =   315
         Left            =   2430
         TabIndex        =   48
         Top             =   1020
         Width           =   1845
         VariousPropertyBits=   746604571
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "3254;556"
         Value           =   "0"
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   450
         TabIndex        =   47
         Top             =   660
         Width           =   1935
         BackColor       =   16777215
         Caption         =   "Commision Percentage1:"
         Size            =   "3413;450"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtComm1 
         Height          =   315
         Left            =   2430
         TabIndex        =   46
         Top             =   600
         Width           =   1845
         VariousPropertyBits=   746604571
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "3254;556"
         Value           =   "0"
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   450
         TabIndex        =   45
         Top             =   240
         Width           =   1935
         BackColor       =   16777215
         Caption         =   "Hourly Wage:"
         Size            =   "3413;450"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtWage 
         Height          =   315
         Left            =   2430
         TabIndex        =   44
         Top             =   180
         Width           =   1845
         VariousPropertyBits=   746604571
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "3254;556"
         Value           =   "0"
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image Image5 
         Height          =   3465
         Left            =   120
         Top             =   1440
         Width           =   4155
         BackColor       =   16777215
         Size            =   "7329;6112"
      End
      Begin MSForms.Image Image4 
         Height          =   5025
         Left            =   0
         Top             =   0
         Width           =   8595
         BorderColor     =   12632256
         BackColor       =   16777215
         Size            =   "15161;8864"
      End
   End
   Begin VB.PictureBox picPage1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   10
      ScaleHeight     =   5025
      ScaleWidth      =   8895
      TabIndex        =   8
      Top             =   1110
      Width           =   8895
      Begin MSForms.CheckBox auReservation 
         Height          =   345
         Left            =   660
         TabIndex        =   19
         Top             =   270
         Width           =   1575
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Reservations"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auRooms 
         Height          =   345
         Left            =   660
         TabIndex        =   18
         Top             =   705
         Width           =   1575
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Rooms"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auGuests 
         Height          =   345
         Left            =   660
         TabIndex        =   17
         Top             =   1125
         Width           =   1575
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Accounts"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auReports 
         Height          =   345
         Left            =   660
         TabIndex        =   16
         Top             =   2850
         Width           =   1575
         VariousPropertyBits=   746588179
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Reports"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auUsers 
         Height          =   345
         Left            =   660
         TabIndex        =   15
         Top             =   2430
         Width           =   1575
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Users"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auSettings 
         Height          =   345
         Left            =   660
         TabIndex        =   14
         Top             =   3285
         Width           =   1575
         VariousPropertyBits=   746588179
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Settings"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auCheckin 
         Height          =   345
         Left            =   660
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Check in "
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auCheckout 
         Height          =   345
         Left            =   660
         TabIndex        =   12
         Top             =   1995
         Width           =   1575
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Check out"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auSales 
         Height          =   345
         Left            =   660
         TabIndex        =   11
         Top             =   3720
         Width           =   1575
         VariousPropertyBits=   746588179
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Sales"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox auInventory 
         Height          =   345
         Left            =   660
         TabIndex        =   10
         Top             =   4155
         Width           =   1575
         VariousPropertyBits=   746588179
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;609"
         Value           =   "0"
         Caption         =   "Inventory"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image Image1 
         Height          =   4995
         Left            =   150
         Top             =   0
         Width           =   8595
         BorderColor     =   12632256
         BackColor       =   16777215
         Size            =   "15161;8811"
         VariousPropertyBits=   19
      End
   End
   Begin MSForms.CheckBox chkApply 
      Height          =   285
      Left            =   150
      TabIndex        =   69
      Tag             =   "1"
      Top             =   6270
      Width           =   3675
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "6482;503"
      Value           =   "0"
      Caption         =   "Apply Commision1 to all Departments"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblNumber 
      Height          =   315
      Left            =   2940
      TabIndex        =   5
      Top             =   4830
      Visible         =   0   'False
      Width           =   1185
      BackColor       =   -2147483643
      Size            =   "2090;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "Log on Name:"
      Size            =   "2461;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblLoginName 
      Height          =   225
      Left            =   1710
      TabIndex        =   3
      Top             =   260
      Width           =   4125
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Size            =   "7276;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image3 
      Height          =   285
      Left            =   1650
      Top             =   210
      Width           =   4875
      BackColor       =   16777215
      Size            =   "8599;503"
   End
End
Attribute VB_Name = "frmUserAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmbDepartments_Change()
    If cmbDepartments.Text = "<Select a Major Department>" Then
        grdSub.Rows = 1
    Else
        grdSub.Rows = 1
        ActiveReadServer2 "Select * from departments where Dept_Parent = '" & Trim(Mid(cmbDepartments.Text, 1, InStrRev(cmbDepartments, "-") - 1) & "' and Dept_Type <>0 order by Department_No")
        While Not rs2.EOF
            grdSub.Rows = grdSub.Rows + 1
            grdSub.TextMatrix(grdSub.Rows - 1, 0) = rs2.Fields("Department_No") & " - " & rs2.Fields("Dept_Name")
            grdSub.TextMatrix(grdSub.Rows - 1, 1) = "0"
            rs2.MoveNext
        Wend
        rs2.Close
    End If
End Sub

Private Sub cmdForms_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
        Case 2
            Select Case auReservation.Value
                Case False: uares = 0
                Case True: uares = 1
            End Select
            Select Case auRooms.Value
                Case False: uarooms = 0
                Case True: uarooms = 1
            End Select
             Select Case auGuests.Value
                Case False: uaguests = 0
                Case True: uaguests = 1
            End Select
             Select Case auCheckin.Value
                Case False: uacheckin = 0
                Case True: uacheckin = 1
            End Select
            Select Case auCheckout.Value
                Case False: uaCheckout = 0
                Case True: uaCheckout = 1
            End Select
              Select Case auUsers.Value
                Case False: uausers = 0
                Case True: uausers = 1
            End Select
              Select Case auReports.Value
                Case False: uaReports = 0
                Case True: uaReports = 1
            End Select
            Select Case auSettings.Value
                Case False: uasettings = 0
                Case True: uasettings = 1
            End Select
            Select Case auSales.Value
                Case False: uaSales = 0
                Case True: uaSales = 1
            End Select
            Select Case auInventory.Value
                Case False: uaInventory = 0
                Case True: uaInventory = 1
            End Select
            Select Case chkBox(0).Value
                Case False: cashS = 0
                Case True: cashS = 1
            End Select
            Select Case chkBox(1).Value
                Case False: chequeS = 0
                Case True: chequeS = 1
            End Select
            Select Case chkBox(2).Value
                Case False: cardS = 0
                Case True: cardS = 1
            End Select
            Select Case chkBox(3).Value
                Case False: chargeS = 0
                Case True: chargeS = 1
            End Select
            Select Case chkBox(4).Value
                Case False: loyaltyS = 0
                Case True: loyaltyS = 1
            End Select
            Select Case chkBox(5).Value
                Case False: itemC = 0
                Case True: itemC = 1
            End Select
            Select Case chkBox(6).Value
                Case False: void = 0
                Case True: void = 1
            End Select
            Select Case chkBox(7).Value
                Case False: rtmd = 0
                Case True: rtmd = 1
            End Select
            Select Case chkBox(8).Value
                Case False: discP = 0
                Case True: discP = 1
            End Select
            Select Case chkBox(9).Value
                Case False: discA = 0
                Case True: discA = 1
            End Select
            Select Case chkBox(10).Value
                Case False: overT = 0
                Case True: overT = 1
            End Select
            Select Case chkBox(11).Value
                Case False: PayO = 0
                Case True: PayO = 1
            End Select
            Select Case chkBox(12).Value
                Case False: pickU = 0
                Case True: pickU = 1
            End Select
            Select Case chkBox(13).Value
                Case False: loan = 0
                Case True: loan = 1
            End Select
            Select Case chkBox(14).Value
                Case False: receiveA = 0
                Case True: receiveA = 1
            End Select
            Select Case chkBox(15).Value
                Case False: storeR = 0
                Case True: storeR = 1
            End Select
            Select Case chkBox(16).Value
                Case False: splitT = 0
                Case True: splitT = 1
            End Select
            Select Case chkBox(17).Value
                Case False: transC = 0
                Case True: transC = 1
            End Select
            Select Case chkBox(18).Value
                Case False: TabTrans = 0
                Case True: TabTrans = 1
            End Select
            Select Case chkBox(25).Value
                Case False: NoSales = 0
                Case True: NoSales = 1
            End Select
            Select Case chkBox(24).Value
                Case False: Ullages = 0
                Case True: Ullages = 1
            End Select
            Select Case chkBox(19).Value
                Case False: AmtOver = 0
                Case True: AmtOver = 1
            End Select
            Select Case chkBox(20).Value
                Case False: bufP = 0
                Case True: bufP = 1
            End Select
            Select Case chkBox(28).Value
                Case False: Reprint = 0
                Case True: Reprint = 1
            End Select
            Select Case chkBox(21).Value
                Case False: pluSc = 0
                Case True: pluSc = 1
            End Select
            Select Case chkBox(22).Value
                Case False: Total_Clear = 0
                Case True: Total_Clear = 1
            End Select
            Select Case chkLog.Value
                Case False: Logg = 0
                Case True: Logg = 1
            End Select
            Select Case chkAllTables.Value
                Case False: All_Tables = 0
                Case True: All_Tables = 1
            End Select
            Select Case chkBox(23).Value
                Case False: App_Exit = 0
                Case True: App_Exit = 1
            End Select
            
             Select Case chkBox(26).Value
                Case False: Service_Charge = 0
                Case True: Service_Charge = 1
            End Select
             Select Case chkBox(27).Value
                Case False: Quotes = 0
                Case True: Quotes = 1
            End Select
            
            Select Case Me.chkBox(29)
               Case False: Ownner_transfer = 0
                Case True: Ownner_transfer = 1
            End Select
            
            Select Case chkPanels(0).Value
                Case False: Draw_Cash = 0
                Case True: Draw_Cash = 1
            End Select
            Select Case chkPanels(1).Value
                Case False: Draw_Card = 0
                Case True: Draw_Card = 1
            End Select
            Select Case chkPanels(2).Value
                Case False: Draw_Cheque = 0
                Case True: Draw_Cheque = 1
            End Select
            Select Case chkPanels(3).Value
                Case False: Draw_Charge = 0
                Case True: Draw_Charge = 1
            End Select
            Select Case chkPanels(4).Value
                Case False: Draw_Loyalty = 0
                Case True: Draw_Loyalty = 1
            End Select
            Select Case chkComms.Value
                Case False: UserRecord.Com_Calc = 0
                Case True: UserRecord.Com_Calc = 1
            End Select
            Select Case chkBarCash.Value
                Case False: UserRecord.Bar_Cash = 0
                Case True: UserRecord.Bar_Cash = 1
            End Select
            Select Case chkBarCash.Value
                Case False: UserRecord.Bar_Cash = 0
                Case True: UserRecord.Bar_Cash = 1
            End Select
            
            ActiveUpdateServer "Update users set " & "Ua_Reservations = " & uares & ",Ua_Rooms=" & uarooms & ",Ua_Guests=" & uaguests & ",Ua_Checkin=" & uacheckin & _
            ",Ua_Checkout=" & uaCheckout & ",Ua_Users=" & uausers & ",Ua_Reports=" & uaReports & ",Ua_Sales=" & uaSales & _
            ",Cash_Sales=" & cashS & ",Cheque_Sales=" & chequeS & _
            ",Card_Sales=" & cardS & ",Charge_Sales=" & chargeS & _
            ",Loyalty_Sales=" & loyaltyS & ",Item_Corrects=" & itemC & _
            ",Voids=" & void & ",Returns=" & rtmd & ",Ullages =" & Ullages & _
            ",Disc_Perc=" & discP & ",Disc_Amt=" & discA & ",Over_Tender=" & overT & ",Payouts=" & PayO & _
            ",Pickups=" & pickU & _
            ",Loans=" & loan & ",No_Sales=" & NoSales & _
            ",Receive_Acc=" & receiveA & ",Split_Tenders=" & splitT & ",Buffer_Print=" & bufP & ",Trans_Store=" & storeR & _
            ",App_Exit=" & App_Exit & ",Trans_Clear=" & transC & ",Transfer=" & TabTrans & _
            ",Reprint=" & Reprint & ",Quotes =" & Quotes & ",Override= " & AmtOver & _
            ",Search=" & pluSc & ", Owner_transfer=" & Ownner_transfer & _
            ",Com_Calc=" & UserRecord.Com_Calc & _
            ",Logged_In=" & Logg & _
            ",Draw_Cash=" & Draw_Cash & _
            ",Draw_Card=" & Draw_Card & _
            ",Draw_Cheque=" & Draw_Cheque & _
            ",Draw_Charge=" & Draw_Charge & _
            ",Draw_Loyalty=" & Draw_Loyalty & _
            ",Service_Charge=" & Service_Charge & _
            ",Bar_Cash=" & UserRecord.Bar_Cash & _
            ",All_Tables=" & All_Tables & ",Wage=" & Val(txtWage.Text) & _
            ",Total_Clear=" & Val(Total_Clear) & ",Comm1=" & Val(txtComm1.Text) & ",Comm2=" & Val(txtComm2.Text) & _
            ",Ua_Inventory=" & uaInventory & ",Ua_settings=" & uasettings & " where user_no =" & lblNumber.Caption
            Unload Me
    End Select
End Sub

Private Sub cmdTab_Click(Index As Integer)
    Select Case Index
        Case 0
            picPage2.Visible = False
            picPage3.Visible = False
            picPage1.Visible = True
            cmdTab(0).FontTextCaption.Bold = True
            cmdTab(1).FontTextCaption.Bold = False
            cmdTab(2).FontTextCaption.Bold = False
        Case 1
            picPage1.Visible = False
            picPage3.Visible = False
            picPage2.Visible = True
            cmdTab(1).FontTextCaption.Bold = True
            cmdTab(0).FontTextCaption.Bold = False
            cmdTab(2).FontTextCaption.Bold = False
        Case 2
            picPage1.Visible = False
            picPage2.Visible = False
            picPage3.Visible = True
            cmdTab(2).FontTextCaption.Bold = True
            cmdTab(0).FontTextCaption.Bold = False
            cmdTab(1).FontTextCaption.Bold = False
            txtWage.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    cmbDepartments.Clear
    ActiveReadServer2 "Select * from Departments where Dept_Type =0 order by Department_No"
    While Not rs2.EOF
        cmbDepartments.AddItem rs2.Fields("Department_No") & " - " & rs2.Fields("Dept_Name")
        rs2.MoveNext
    Wend
    rs2.Close
    cmbDepartments.AddItem "<Select a Major Department>"
    cmbDepartments.Text = "<Select a Major Department>"
    grdSub.TextMatrix(0, 0) = "Department"
    grdSub.TextMatrix(0, 1) = "Comm%"
    grdSub.ColWidth(0) = grdSub.Width * 0.7
    grdSub.ColWidth(1) = grdSub.Width * 0.2
End Sub

Private Sub grdSub_Click()
    If grdSub.Col = 0 Then grdSub.Col = 1
End Sub
Private Sub grdSub_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 45, 48 To 57, 96 To 105, 109, 110, 189
            Select Case grdSub.Col
                Case 1
                    grdSub.EditCell
            End Select
        Case 1
    End Select
End Sub

Private Sub grdSub_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(grdSub.EditText, ".") <> 0 And KeyAscii = 46 Then
        KeyAscii = 0
    End If
    Select Case KeyAscii
        Case 8, 13, 27, 45, 46, 48 To 57
        Case Else
            If Col = 1 Then KeyAscii = 0
    End Select
End Sub

Private Sub txtComm1_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtComm1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtComm2_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtComm2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtWage_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtWage_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtWage_LostFocus()
    txtWage.Text = Format(txtWage.Text, "0.00")
End Sub
