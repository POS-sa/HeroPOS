VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmHappy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Happy Hour"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   Icon            =   "frmHappy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Happy Hour Days"
      ForeColor       =   &H80000008&
      Height          =   3345
      Index           =   1
      Left            =   3600
      TabIndex        =   22
      Top             =   3870
      Width           =   2445
      Begin btButtonEx.ButtonEx cmdSave 
         Height          =   345
         Index           =   1
         Left            =   1320
         TabIndex        =   23
         Top             =   2910
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Appearance      =   3
         Caption         =   "Save"
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
      Begin MSForms.CheckBox chkWeek1 
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   30
         Top             =   360
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Monday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek1 
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   29
         Top             =   765
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Tuesday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek1 
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   28
         Top             =   1170
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Wednesday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek1 
         Height          =   285
         Index           =   3
         Left            =   210
         TabIndex        =   27
         Top             =   1575
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Thursday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek1 
         Height          =   285
         Index           =   4
         Left            =   210
         TabIndex        =   26
         Top             =   1980
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Friday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek1 
         Height          =   285
         Index           =   5
         Left            =   210
         TabIndex        =   25
         Top             =   2385
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Saturday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek1 
         Height          =   285
         Index           =   6
         Left            =   210
         TabIndex        =   24
         Top             =   2790
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Sunday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.ComboBox cmbHappy 
      Height          =   315
      Index           =   1
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4800
      Width           =   1845
   End
   Begin VB.ComboBox cmbHappy 
      Height          =   315
      Index           =   0
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1140
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Happy Hour Days"
      ForeColor       =   &H80000008&
      Height          =   3345
      Index           =   0
      Left            =   3600
      TabIndex        =   8
      Top             =   210
      Width           =   2445
      Begin btButtonEx.ButtonEx cmdSave 
         Height          =   345
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   2910
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Appearance      =   3
         Caption         =   "Save"
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
      Begin MSForms.CheckBox chkWeek 
         Height          =   285
         Index           =   6
         Left            =   210
         TabIndex        =   15
         Top             =   2790
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Sunday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek 
         Height          =   285
         Index           =   5
         Left            =   210
         TabIndex        =   14
         Top             =   2385
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Saturday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek 
         Height          =   285
         Index           =   4
         Left            =   210
         TabIndex        =   13
         Top             =   1980
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Friday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek 
         Height          =   285
         Index           =   3
         Left            =   210
         TabIndex        =   12
         Top             =   1575
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Thursday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek 
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   11
         Top             =   1170
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Wednesday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek 
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Top             =   765
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Tuesday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkWeek 
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Top             =   360
         Width           =   1605
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2831;503"
         Value           =   "0"
         Caption         =   "Monday"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComCtl2.DTPicker DTStart 
      Height          =   345
      Index           =   0
      Left            =   1620
      TabIndex        =   6
      Top             =   300
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      Format          =   62324738
      CurrentDate     =   38919
   End
   Begin btButtonEx.ButtonEx cmdStart 
      Height          =   525
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   3000
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   926
      Appearance      =   3
      Caption         =   "Start Now"
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
   Begin btButtonEx.ButtonEx cmdStop 
      Height          =   525
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   3000
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   926
      Appearance      =   3
      Caption         =   "Stop Now"
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
   Begin MSComCtl2.DTPicker DTStop 
      Height          =   345
      Index           =   0
      Left            =   1620
      TabIndex        =   7
      Top             =   720
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      Format          =   62324738
      CurrentDate     =   38919
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   525
      Index           =   0
      Left            =   2430
      TabIndex        =   17
      Top             =   3000
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   926
      Appearance      =   3
      Caption         =   "Exit"
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
   Begin BTNENHLib4.BtnEnh fmType 
      Height          =   1005
      Index           =   0
      Left            =   210
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3285
      _Version        =   524298
      _ExtentX        =   5794
      _ExtentY        =   1773
      _StockProps     =   66
      Caption         =   "Provisional Booking"
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
      Shape           =   1
      CornerFactor    =   10
      Surface         =   3
      PictureTranspColor=   192
      BackColorContainer=   14215660
      ShadowColor     =   16777215
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      ForeColorDisabled=   12640511
      UserData        =   0.1
      textCaption     =   "frmHappy.frx":000C
      textLT          =   "frmHappy.frx":0092
      textCT          =   "frmHappy.frx":00AA
      textRT          =   "frmHappy.frx":00C2
      textLM          =   "frmHappy.frx":00DA
      textRM          =   "frmHappy.frx":00F2
      textLB          =   "frmHappy.frx":010A
      textCB          =   "frmHappy.frx":0122
      textRB          =   "frmHappy.frx":013A
      colorBack       =   "frmHappy.frx":0152
      colorIntern     =   "frmHappy.frx":017C
      colorMO         =   "frmHappy.frx":01A6
      colorFocus      =   "frmHappy.frx":01D0
      colorDisabled   =   "frmHappy.frx":01FA
      colorPressed    =   "frmHappy.frx":0224
      HollowFrame     =   -1  'True
      LightDirection  =   8
   End
   Begin MSComCtl2.DTPicker DTStart 
      Height          =   345
      Index           =   1
      Left            =   1620
      TabIndex        =   31
      Top             =   3960
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      Format          =   62324738
      CurrentDate     =   38919
   End
   Begin btButtonEx.ButtonEx cmdStart 
      Height          =   525
      Index           =   1
      Left            =   210
      TabIndex        =   32
      Top             =   6660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   926
      Appearance      =   3
      Caption         =   "Start Now"
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
   Begin btButtonEx.ButtonEx cmdStop 
      Height          =   525
      Index           =   1
      Left            =   1320
      TabIndex        =   33
      Top             =   6660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   926
      Appearance      =   3
      Caption         =   "Stop Now"
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
   Begin MSComCtl2.DTPicker DTStop 
      Height          =   345
      Index           =   1
      Left            =   1620
      TabIndex        =   34
      Top             =   4380
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      Format          =   62324738
      CurrentDate     =   38919
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   525
      Index           =   1
      Left            =   2430
      TabIndex        =   35
      Top             =   6660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   926
      Appearance      =   3
      Caption         =   "Exit"
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
   Begin BTNENHLib4.BtnEnh fmType 
      Height          =   1005
      Index           =   1
      Left            =   210
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5550
      Width           =   3285
      _Version        =   524298
      _ExtentX        =   5794
      _ExtentY        =   1773
      _StockProps     =   66
      Caption         =   "Provisional Booking"
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
      Shape           =   1
      CornerFactor    =   10
      Surface         =   3
      PictureTranspColor=   192
      BackColorContainer=   14215660
      ShadowColor     =   16777215
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      ForeColorDisabled=   12640511
      UserData        =   0.1
      textCaption     =   "frmHappy.frx":024E
      textLT          =   "frmHappy.frx":02D4
      textCT          =   "frmHappy.frx":02EC
      textRT          =   "frmHappy.frx":0304
      textLM          =   "frmHappy.frx":031C
      textRM          =   "frmHappy.frx":0334
      textLB          =   "frmHappy.frx":034C
      textCB          =   "frmHappy.frx":0364
      textRB          =   "frmHappy.frx":037C
      colorBack       =   "frmHappy.frx":0394
      colorIntern     =   "frmHappy.frx":03BE
      colorMO         =   "frmHappy.frx":03E8
      colorFocus      =   "frmHappy.frx":0412
      colorDisabled   =   "frmHappy.frx":043C
      colorPressed    =   "frmHappy.frx":0466
      HollowFrame     =   -1  'True
      LightDirection  =   8
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Happy Hour No.2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   41
      Top             =   3690
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Starts at:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   40
      Top             =   3990
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Stops at:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   39
      Top             =   4425
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Use Selling Price:"
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   38
      Top             =   4860
      Width           =   1335
   End
   Begin MSForms.CheckBox chkStart 
      Height          =   315
      Index           =   1
      Left            =   1590
      TabIndex        =   37
      Top             =   5160
      Width           =   1755
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3096;556"
      Value           =   "0"
      Caption         =   "Auto Switch "
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Happy Hour No.1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   30
      Width           =   1335
   End
   Begin MSForms.CheckBox chkStart 
      Height          =   315
      Index           =   0
      Left            =   1590
      TabIndex        =   3
      Top             =   1500
      Width           =   1755
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3096;556"
      Value           =   "0"
      Caption         =   "Auto Switch "
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Use Selling Price:"
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Stops at:"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   765
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Starts at:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   330
      Width           =   1335
   End
   Begin MSForms.Image Image1 
      DragMode        =   1  'Automatic
      Height          =   3525
      Index           =   0
      Left            =   120
      Top             =   150
      Width           =   6075
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "10716;6218"
   End
   Begin MSForms.Image Image1 
      DragMode        =   1  'Automatic
      Height          =   3525
      Index           =   1
      Left            =   120
      Top             =   3810
      Width           =   6075
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "10716;6218"
   End
End
Attribute VB_Name = "frmHappy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSave_Click(Index As Integer)
    If Index = 0 Then
        ActiveUpdateServer "Delete from Happy_Hour"
        Select Case cmbHappy(0).Text
            Case "Price One"
                PriceNo = 1
            Case "Price Two"
                PriceNo = 2
            Case "Price Three"
                PriceNo = 3
            Case "Price Four"
                PriceNo = 4
            Case "Price Five"
                PriceNo = 5
            Case "Price Six"
                PriceNo = 6
        End Select
        For i = 1 To 7
            ActiveUpdateServer "INSERT INTO [Happy_Hour]([Week_Day], [Start_Time], [Stop_Time], [Selling_Price],[Auto_Change],[Active])" & _
            "VALUES(" & i & ", '" & DTStart(0).Value & "', '" & DTStop(0).Value & "', " & PriceNo & "," & Abs(chkStart(0).Value) & "," & Abs(chkWeek(i - 1).Value) & ")"
        Next i
    End If
    If Index = 1 Then
        ActiveUpdateServer "Delete from Happy_Hour1"
        Select Case cmbHappy(1).Text
            Case "Price One"
                PriceNo = 1
            Case "Price Two"
                PriceNo = 2
            Case "Price Three"
                PriceNo = 3
            Case "Price Four"
                PriceNo = 4
            Case "Price Five"
                PriceNo = 5
            Case "Price Six"
                PriceNo = 6
        End Select
        For i = 1 To 7
            ActiveUpdateServer "INSERT INTO [Happy_Hour1]([Week_Day], [Start_Time], [Stop_Time], [Selling_Price],[Auto_Change],[Active])" & _
            "VALUES(" & i & ", '" & DTStart(1).Value & "', '" & DTStop(1).Value & "', " & PriceNo & "," & Abs(chkStart(1).Value) & "," & Abs(chkWeek1(i - 1).Value) & ")"
        Next i
    End If
    Unload Me
End Sub

Private Sub ButtonEx2_Click(Index As Integer)
    Unload Me
End Sub
Private Sub cmdStart_Click(Index As Integer)
    If Index = 0 Then
        ActiveUpdateServer "Update Branch_Details set Happy_Active = 1"
        ActiveReadServer "Select Selling_Price from Happy_Hour"
        If rs.RecordCount > 0 Then
            HappyHourPrice = rs.Fields("Selling_Price")
        End If
        rs.Close
        HappyHour = 1
    End If
    If Index = 1 Then
        ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 1"
        ActiveReadServer "Select Selling_Price from Happy_Hour1"
        If rs.RecordCount > 0 Then
            HappyHourPrice1 = rs.Fields("Selling_Price")
        End If
        rs.Close
        HappyHour1 = 1
    End If
    Unload Me
End Sub
Private Sub cmdStop_Click(Index As Integer)
    If Index = 0 Then
        ActiveUpdateServer "Update Branch_Details set Happy_Active = 0"
        HappyHour = 0
    End If
    If Index = 1 Then
        ActiveUpdateServer "Update Branch_Details set Happy_Active1 = 0"
        HappyHour1 = 0
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    Select Case HappyHour
        Case 0
            fmType(0).Caption = "Happy Hour In-Active"
        Case 1
            fmType(0).Caption = "Happy Hour Active"
    End Select
    i = 0
    ActiveReadServer "Select * from Happy_Hour order by Week_Day"
    While Not rs.EOF
        DTStart(0).Value = rs.Fields("Start_Time")
        DTStop(0).Value = rs.Fields("Stop_Time")
        chkWeek(i).Value = rs.Fields("Active")
        chkStart(0).Value = rs.Fields("Auto_Change")
        Select Case rs.Fields("Selling_Price")
            Case 1
                cmbHappy(0).ListIndex = 0
            Case 2
                cmbHappy(0).ListIndex = 1
            Case 3
                cmbHappy(0).ListIndex = 2
            Case 4
                cmbHappy(0).ListIndex = 3
            Case 5
                cmbHappy(0).ListIndex = 4
            Case 6
                cmbHappy(0).ListIndex = 5
        End Select
        rs.MoveNext
        i = i + 1
    Wend
    rs.Close
    
    Select Case HappyHour1
        Case 0
            fmType(1).Caption = "Happy Hour In-Active"
        Case 1
            fmType(1).Caption = "Happy Hour Active"
    End Select
    i = 0
    ActiveReadServer "Select * from Happy_Hour1 order by Week_Day"
    While Not rs.EOF
        DTStart(1).Value = rs.Fields("Start_Time")
        DTStop(1).Value = rs.Fields("Stop_Time")
        chkWeek1(i).Value = rs.Fields("Active")
        chkStart(1).Value = rs.Fields("Auto_Change")
        Select Case rs.Fields("Selling_Price")
            Case 1
                cmbHappy(1).ListIndex = 0
            Case 2
                cmbHappy(1).ListIndex = 1
            Case 3
                cmbHappy(1).ListIndex = 2
            Case 4
                cmbHappy(1).ListIndex = 3
            Case 5
                cmbHappy(1).ListIndex = 4
            Case 6
                cmbHappy(1).ListIndex = 5
        End Select
        rs.MoveNext
        i = i + 1
    Wend
    rs.Close
End Sub

Private Sub Form_Load()
    cmbHappy(0).AddItem "Price One"
    cmbHappy(0).AddItem "Price Two"
    cmbHappy(0).AddItem "Price Three"
    cmbHappy(0).AddItem "Price Four"
    cmbHappy(0).AddItem "Price Five"
    cmbHappy(0).AddItem "Price Six"
    cmbHappy(1).AddItem "Price One"
    cmbHappy(1).AddItem "Price Two"
    cmbHappy(1).AddItem "Price Three"
    cmbHappy(1).AddItem "Price Four"
    cmbHappy(1).AddItem "Price Five"
    cmbHappy(1).AddItem "Price Six"
    
    ActiveReadServer "Select * from Happy_Hour"
    If rs.RecordCount = 0 Then
        For i = 1 To 7
            ActiveUpdateServer "INSERT INTO [Happy_Hour]([Week_Day], [Start_Time], [Stop_Time], [Selling_Price],[Auto_Change],[Active])" & _
            "VALUES(" & i & ", '" & DTStart(0).Value & "', '" & DTStop(0).Value & "', 1,0,0)"
        Next i
    End If
    rs.Close
    cmbHappy(0).Text = "Price One"
    
    ActiveReadServer "Select * from Happy_Hour1"
    If rs.RecordCount = 0 Then
        For i = 1 To 7
            ActiveUpdateServer "INSERT INTO [Happy_Hour1]([Week_Day], [Start_Time], [Stop_Time], [Selling_Price],[Auto_Change],[Active])" & _
            "VALUES(" & i & ", '" & DTStart(1).Value & "', '" & DTStop(1).Value & "', 1,0,0)"
        Next i
    End If
    rs.Close
    cmbHappy(1).Text = "Price One"
End Sub
