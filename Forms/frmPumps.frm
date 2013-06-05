VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmPumps 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   10410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   11
      Left            =   7830
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   176
      Top             =   4950
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   11
         Left            =   0
         TabIndex        =   177
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No12"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":0000
         textLT          =   "frmPumps.frx":0072
         textCT          =   "frmPumps.frx":008A
         textRT          =   "frmPumps.frx":00A2
         textLM          =   "frmPumps.frx":00BA
         textRM          =   "frmPumps.frx":00D2
         textLB          =   "frmPumps.frx":00EA
         textCB          =   "frmPumps.frx":0102
         textRB          =   "frmPumps.frx":011A
         colorBack       =   "frmPumps.frx":0132
         colorIntern     =   "frmPumps.frx":015C
         colorMO         =   "frmPumps.frx":0186
         colorFocus      =   "frmPumps.frx":01B0
         colorDisabled   =   "frmPumps.frx":01DA
         colorPressed    =   "frmPumps.frx":0204
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   22
         Left            =   0
         TabIndex        =   178
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":022E
         textLT          =   "frmPumps.frx":029E
         textCT          =   "frmPumps.frx":02B6
         textRT          =   "frmPumps.frx":02CE
         textLM          =   "frmPumps.frx":02E6
         textRM          =   "frmPumps.frx":02FE
         textLB          =   "frmPumps.frx":0316
         textCB          =   "frmPumps.frx":032E
         textRB          =   "frmPumps.frx":0346
         colorBack       =   "frmPumps.frx":035E
         colorIntern     =   "frmPumps.frx":0388
         colorMO         =   "frmPumps.frx":03B2
         colorFocus      =   "frmPumps.frx":03DC
         colorDisabled   =   "frmPumps.frx":0406
         colorPressed    =   "frmPumps.frx":0430
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   23
         Left            =   1230
         TabIndex        =   179
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":045A
         textLT          =   "frmPumps.frx":04CA
         textCT          =   "frmPumps.frx":04E2
         textRT          =   "frmPumps.frx":04FA
         textLM          =   "frmPumps.frx":0512
         textRM          =   "frmPumps.frx":052A
         textLB          =   "frmPumps.frx":0542
         textCB          =   "frmPumps.frx":055A
         textRB          =   "frmPumps.frx":0572
         colorBack       =   "frmPumps.frx":058A
         colorIntern     =   "frmPumps.frx":05B4
         colorMO         =   "frmPumps.frx":05DE
         colorFocus      =   "frmPumps.frx":0608
         colorDisabled   =   "frmPumps.frx":0632
         colorPressed    =   "frmPumps.frx":065C
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   23
         Left            =   1380
         TabIndex        =   191
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   23
         Left            =   1380
         TabIndex        =   190
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   23
         Left            =   1380
         TabIndex        =   189
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   23
         Left            =   1350
         TabIndex        =   188
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   23
         Left            =   1350
         TabIndex        =   187
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   23
         Left            =   1350
         TabIndex        =   186
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   22
         Left            =   60
         TabIndex        =   185
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   22
         Left            =   60
         TabIndex        =   184
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   22
         Left            =   60
         TabIndex        =   183
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   22
         Left            =   90
         TabIndex        =   182
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   22
         Left            =   90
         TabIndex        =   181
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   22
         Left            =   90
         TabIndex        =   180
         Top             =   840
         Width           =   525
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   11
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   10
      Left            =   7830
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   160
      Top             =   2490
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   10
         Left            =   0
         TabIndex        =   161
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No8"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":0686
         textLT          =   "frmPumps.frx":06F6
         textCT          =   "frmPumps.frx":070E
         textRT          =   "frmPumps.frx":0726
         textLM          =   "frmPumps.frx":073E
         textRM          =   "frmPumps.frx":0756
         textLB          =   "frmPumps.frx":076E
         textCB          =   "frmPumps.frx":0786
         textRB          =   "frmPumps.frx":079E
         colorBack       =   "frmPumps.frx":07B6
         colorIntern     =   "frmPumps.frx":07E0
         colorMO         =   "frmPumps.frx":080A
         colorFocus      =   "frmPumps.frx":0834
         colorDisabled   =   "frmPumps.frx":085E
         colorPressed    =   "frmPumps.frx":0888
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   20
         Left            =   0
         TabIndex        =   162
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":08B2
         textLT          =   "frmPumps.frx":0922
         textCT          =   "frmPumps.frx":093A
         textRT          =   "frmPumps.frx":0952
         textLM          =   "frmPumps.frx":096A
         textRM          =   "frmPumps.frx":0982
         textLB          =   "frmPumps.frx":099A
         textCB          =   "frmPumps.frx":09B2
         textRB          =   "frmPumps.frx":09CA
         colorBack       =   "frmPumps.frx":09E2
         colorIntern     =   "frmPumps.frx":0A0C
         colorMO         =   "frmPumps.frx":0A36
         colorFocus      =   "frmPumps.frx":0A60
         colorDisabled   =   "frmPumps.frx":0A8A
         colorPressed    =   "frmPumps.frx":0AB4
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   21
         Left            =   1230
         TabIndex        =   163
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":0ADE
         textLT          =   "frmPumps.frx":0B4E
         textCT          =   "frmPumps.frx":0B66
         textRT          =   "frmPumps.frx":0B7E
         textLM          =   "frmPumps.frx":0B96
         textRM          =   "frmPumps.frx":0BAE
         textLB          =   "frmPumps.frx":0BC6
         textCB          =   "frmPumps.frx":0BDE
         textRB          =   "frmPumps.frx":0BF6
         colorBack       =   "frmPumps.frx":0C0E
         colorIntern     =   "frmPumps.frx":0C38
         colorMO         =   "frmPumps.frx":0C62
         colorFocus      =   "frmPumps.frx":0C8C
         colorDisabled   =   "frmPumps.frx":0CB6
         colorPressed    =   "frmPumps.frx":0CE0
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   21
         Left            =   1380
         TabIndex        =   175
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   21
         Left            =   1380
         TabIndex        =   174
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   21
         Left            =   1380
         TabIndex        =   173
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   21
         Left            =   1350
         TabIndex        =   172
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   21
         Left            =   1350
         TabIndex        =   171
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   21
         Left            =   1350
         TabIndex        =   170
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   20
         Left            =   60
         TabIndex        =   169
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   20
         Left            =   60
         TabIndex        =   168
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   20
         Left            =   60
         TabIndex        =   167
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   20
         Left            =   90
         TabIndex        =   166
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   20
         Left            =   90
         TabIndex        =   165
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   20
         Left            =   90
         TabIndex        =   164
         Top             =   840
         Width           =   525
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   10
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   9
      Left            =   7830
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   144
      Top             =   30
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   9
         Left            =   0
         TabIndex        =   145
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No4"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":0D0A
         textLT          =   "frmPumps.frx":0D7A
         textCT          =   "frmPumps.frx":0D92
         textRT          =   "frmPumps.frx":0DAA
         textLM          =   "frmPumps.frx":0DC2
         textRM          =   "frmPumps.frx":0DDA
         textLB          =   "frmPumps.frx":0DF2
         textCB          =   "frmPumps.frx":0E0A
         textRB          =   "frmPumps.frx":0E22
         colorBack       =   "frmPumps.frx":0E3A
         colorIntern     =   "frmPumps.frx":0E64
         colorMO         =   "frmPumps.frx":0E8E
         colorFocus      =   "frmPumps.frx":0EB8
         colorDisabled   =   "frmPumps.frx":0EE2
         colorPressed    =   "frmPumps.frx":0F0C
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   18
         Left            =   0
         TabIndex        =   146
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":0F36
         textLT          =   "frmPumps.frx":0FA6
         textCT          =   "frmPumps.frx":0FBE
         textRT          =   "frmPumps.frx":0FD6
         textLM          =   "frmPumps.frx":0FEE
         textRM          =   "frmPumps.frx":1006
         textLB          =   "frmPumps.frx":101E
         textCB          =   "frmPumps.frx":1036
         textRB          =   "frmPumps.frx":104E
         colorBack       =   "frmPumps.frx":1066
         colorIntern     =   "frmPumps.frx":1090
         colorMO         =   "frmPumps.frx":10BA
         colorFocus      =   "frmPumps.frx":10E4
         colorDisabled   =   "frmPumps.frx":110E
         colorPressed    =   "frmPumps.frx":1138
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   19
         Left            =   1230
         TabIndex        =   147
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":1162
         textLT          =   "frmPumps.frx":11D2
         textCT          =   "frmPumps.frx":11EA
         textRT          =   "frmPumps.frx":1202
         textLM          =   "frmPumps.frx":121A
         textRM          =   "frmPumps.frx":1232
         textLB          =   "frmPumps.frx":124A
         textCB          =   "frmPumps.frx":1262
         textRB          =   "frmPumps.frx":127A
         colorBack       =   "frmPumps.frx":1292
         colorIntern     =   "frmPumps.frx":12BC
         colorMO         =   "frmPumps.frx":12E6
         colorFocus      =   "frmPumps.frx":1310
         colorDisabled   =   "frmPumps.frx":133A
         colorPressed    =   "frmPumps.frx":1364
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   9
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   19
         Left            =   90
         TabIndex        =   159
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   19
         Left            =   90
         TabIndex        =   158
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   19
         Left            =   90
         TabIndex        =   157
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   19
         Left            =   60
         TabIndex        =   156
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   19
         Left            =   60
         TabIndex        =   155
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   19
         Left            =   60
         TabIndex        =   154
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   18
         Left            =   1350
         TabIndex        =   153
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   18
         Left            =   1350
         TabIndex        =   152
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   18
         Left            =   1350
         TabIndex        =   151
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   18
         Left            =   1380
         TabIndex        =   150
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   18
         Left            =   1380
         TabIndex        =   149
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   18
         Left            =   1380
         TabIndex        =   148
         Top             =   840
         Width           =   525
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   8
      Left            =   5250
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   128
      Top             =   4950
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   8
         Left            =   0
         TabIndex        =   129
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No11"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":138E
         textLT          =   "frmPumps.frx":1400
         textCT          =   "frmPumps.frx":1418
         textRT          =   "frmPumps.frx":1430
         textLM          =   "frmPumps.frx":1448
         textRM          =   "frmPumps.frx":1460
         textLB          =   "frmPumps.frx":1478
         textCB          =   "frmPumps.frx":1490
         textRB          =   "frmPumps.frx":14A8
         colorBack       =   "frmPumps.frx":14C0
         colorIntern     =   "frmPumps.frx":14EA
         colorMO         =   "frmPumps.frx":1514
         colorFocus      =   "frmPumps.frx":153E
         colorDisabled   =   "frmPumps.frx":1568
         colorPressed    =   "frmPumps.frx":1592
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   16
         Left            =   0
         TabIndex        =   130
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":15BC
         textLT          =   "frmPumps.frx":162C
         textCT          =   "frmPumps.frx":1644
         textRT          =   "frmPumps.frx":165C
         textLM          =   "frmPumps.frx":1674
         textRM          =   "frmPumps.frx":168C
         textLB          =   "frmPumps.frx":16A4
         textCB          =   "frmPumps.frx":16BC
         textRB          =   "frmPumps.frx":16D4
         colorBack       =   "frmPumps.frx":16EC
         colorIntern     =   "frmPumps.frx":1716
         colorMO         =   "frmPumps.frx":1740
         colorFocus      =   "frmPumps.frx":176A
         colorDisabled   =   "frmPumps.frx":1794
         colorPressed    =   "frmPumps.frx":17BE
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   17
         Left            =   1230
         TabIndex        =   131
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":17E8
         textLT          =   "frmPumps.frx":1858
         textCT          =   "frmPumps.frx":1870
         textRT          =   "frmPumps.frx":1888
         textLM          =   "frmPumps.frx":18A0
         textRM          =   "frmPumps.frx":18B8
         textLB          =   "frmPumps.frx":18D0
         textCB          =   "frmPumps.frx":18E8
         textRB          =   "frmPumps.frx":1900
         colorBack       =   "frmPumps.frx":1918
         colorIntern     =   "frmPumps.frx":1942
         colorMO         =   "frmPumps.frx":196C
         colorFocus      =   "frmPumps.frx":1996
         colorDisabled   =   "frmPumps.frx":19C0
         colorPressed    =   "frmPumps.frx":19EA
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   17
         Left            =   1380
         TabIndex        =   143
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   17
         Left            =   1380
         TabIndex        =   142
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   17
         Left            =   1380
         TabIndex        =   141
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   17
         Left            =   1350
         TabIndex        =   140
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   17
         Left            =   1350
         TabIndex        =   139
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   17
         Left            =   1350
         TabIndex        =   138
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   16
         Left            =   60
         TabIndex        =   137
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   16
         Left            =   60
         TabIndex        =   136
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   16
         Left            =   60
         TabIndex        =   135
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   16
         Left            =   90
         TabIndex        =   134
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   16
         Left            =   90
         TabIndex        =   133
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   16
         Left            =   90
         TabIndex        =   132
         Top             =   840
         Width           =   525
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   8
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   7
      Left            =   5250
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   112
      Top             =   2490
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   7
         Left            =   0
         TabIndex        =   113
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No7"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":1A14
         textLT          =   "frmPumps.frx":1A84
         textCT          =   "frmPumps.frx":1A9C
         textRT          =   "frmPumps.frx":1AB4
         textLM          =   "frmPumps.frx":1ACC
         textRM          =   "frmPumps.frx":1AE4
         textLB          =   "frmPumps.frx":1AFC
         textCB          =   "frmPumps.frx":1B14
         textRB          =   "frmPumps.frx":1B2C
         colorBack       =   "frmPumps.frx":1B44
         colorIntern     =   "frmPumps.frx":1B6E
         colorMO         =   "frmPumps.frx":1B98
         colorFocus      =   "frmPumps.frx":1BC2
         colorDisabled   =   "frmPumps.frx":1BEC
         colorPressed    =   "frmPumps.frx":1C16
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   14
         Left            =   0
         TabIndex        =   114
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":1C40
         textLT          =   "frmPumps.frx":1CB0
         textCT          =   "frmPumps.frx":1CC8
         textRT          =   "frmPumps.frx":1CE0
         textLM          =   "frmPumps.frx":1CF8
         textRM          =   "frmPumps.frx":1D10
         textLB          =   "frmPumps.frx":1D28
         textCB          =   "frmPumps.frx":1D40
         textRB          =   "frmPumps.frx":1D58
         colorBack       =   "frmPumps.frx":1D70
         colorIntern     =   "frmPumps.frx":1D9A
         colorMO         =   "frmPumps.frx":1DC4
         colorFocus      =   "frmPumps.frx":1DEE
         colorDisabled   =   "frmPumps.frx":1E18
         colorPressed    =   "frmPumps.frx":1E42
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   15
         Left            =   1230
         TabIndex        =   115
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":1E6C
         textLT          =   "frmPumps.frx":1EDC
         textCT          =   "frmPumps.frx":1EF4
         textRT          =   "frmPumps.frx":1F0C
         textLM          =   "frmPumps.frx":1F24
         textRM          =   "frmPumps.frx":1F3C
         textLB          =   "frmPumps.frx":1F54
         textCB          =   "frmPumps.frx":1F6C
         textRB          =   "frmPumps.frx":1F84
         colorBack       =   "frmPumps.frx":1F9C
         colorIntern     =   "frmPumps.frx":1FC6
         colorMO         =   "frmPumps.frx":1FF0
         colorFocus      =   "frmPumps.frx":201A
         colorDisabled   =   "frmPumps.frx":2044
         colorPressed    =   "frmPumps.frx":206E
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   15
         Left            =   1380
         TabIndex        =   127
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   15
         Left            =   1380
         TabIndex        =   126
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   15
         Left            =   1380
         TabIndex        =   125
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   15
         Left            =   1350
         TabIndex        =   124
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   15
         Left            =   1350
         TabIndex        =   123
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   15
         Left            =   1350
         TabIndex        =   122
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   14
         Left            =   60
         TabIndex        =   121
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   14
         Left            =   60
         TabIndex        =   120
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   14
         Left            =   60
         TabIndex        =   119
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   14
         Left            =   90
         TabIndex        =   118
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   14
         Left            =   90
         TabIndex        =   117
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   14
         Left            =   90
         TabIndex        =   116
         Top             =   840
         Width           =   525
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   7
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   6
      Left            =   5250
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   96
      Top             =   30
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   6
         Left            =   0
         TabIndex        =   97
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No3"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":2098
         textLT          =   "frmPumps.frx":2108
         textCT          =   "frmPumps.frx":2120
         textRT          =   "frmPumps.frx":2138
         textLM          =   "frmPumps.frx":2150
         textRM          =   "frmPumps.frx":2168
         textLB          =   "frmPumps.frx":2180
         textCB          =   "frmPumps.frx":2198
         textRB          =   "frmPumps.frx":21B0
         colorBack       =   "frmPumps.frx":21C8
         colorIntern     =   "frmPumps.frx":21F2
         colorMO         =   "frmPumps.frx":221C
         colorFocus      =   "frmPumps.frx":2246
         colorDisabled   =   "frmPumps.frx":2270
         colorPressed    =   "frmPumps.frx":229A
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   12
         Left            =   0
         TabIndex        =   98
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":22C4
         textLT          =   "frmPumps.frx":2334
         textCT          =   "frmPumps.frx":234C
         textRT          =   "frmPumps.frx":2364
         textLM          =   "frmPumps.frx":237C
         textRM          =   "frmPumps.frx":2394
         textLB          =   "frmPumps.frx":23AC
         textCB          =   "frmPumps.frx":23C4
         textRB          =   "frmPumps.frx":23DC
         colorBack       =   "frmPumps.frx":23F4
         colorIntern     =   "frmPumps.frx":241E
         colorMO         =   "frmPumps.frx":2448
         colorFocus      =   "frmPumps.frx":2472
         colorDisabled   =   "frmPumps.frx":249C
         colorPressed    =   "frmPumps.frx":24C6
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   13
         Left            =   1230
         TabIndex        =   99
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":24F0
         textLT          =   "frmPumps.frx":2560
         textCT          =   "frmPumps.frx":2578
         textRT          =   "frmPumps.frx":2590
         textLM          =   "frmPumps.frx":25A8
         textRM          =   "frmPumps.frx":25C0
         textLB          =   "frmPumps.frx":25D8
         textCB          =   "frmPumps.frx":25F0
         textRB          =   "frmPumps.frx":2608
         colorBack       =   "frmPumps.frx":2620
         colorIntern     =   "frmPumps.frx":264A
         colorMO         =   "frmPumps.frx":2674
         colorFocus      =   "frmPumps.frx":269E
         colorDisabled   =   "frmPumps.frx":26C8
         colorPressed    =   "frmPumps.frx":26F2
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   6
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   13
         Left            =   90
         TabIndex        =   111
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   13
         Left            =   90
         TabIndex        =   110
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   13
         Left            =   90
         TabIndex        =   109
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   13
         Left            =   60
         TabIndex        =   108
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   13
         Left            =   60
         TabIndex        =   107
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   13
         Left            =   60
         TabIndex        =   106
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   12
         Left            =   1350
         TabIndex        =   105
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   12
         Left            =   1350
         TabIndex        =   104
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   12
         Left            =   1350
         TabIndex        =   103
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   12
         Left            =   1380
         TabIndex        =   102
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   12
         Left            =   1380
         TabIndex        =   101
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   12
         Left            =   1380
         TabIndex        =   100
         Top             =   840
         Width           =   525
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   5
      Left            =   2640
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   80
      Top             =   4950
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   5
         Left            =   0
         TabIndex        =   81
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No10"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":271C
         textLT          =   "frmPumps.frx":278E
         textCT          =   "frmPumps.frx":27A6
         textRT          =   "frmPumps.frx":27BE
         textLM          =   "frmPumps.frx":27D6
         textRM          =   "frmPumps.frx":27EE
         textLB          =   "frmPumps.frx":2806
         textCB          =   "frmPumps.frx":281E
         textRB          =   "frmPumps.frx":2836
         colorBack       =   "frmPumps.frx":284E
         colorIntern     =   "frmPumps.frx":2878
         colorMO         =   "frmPumps.frx":28A2
         colorFocus      =   "frmPumps.frx":28CC
         colorDisabled   =   "frmPumps.frx":28F6
         colorPressed    =   "frmPumps.frx":2920
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   10
         Left            =   0
         TabIndex        =   82
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":294A
         textLT          =   "frmPumps.frx":29BA
         textCT          =   "frmPumps.frx":29D2
         textRT          =   "frmPumps.frx":29EA
         textLM          =   "frmPumps.frx":2A02
         textRM          =   "frmPumps.frx":2A1A
         textLB          =   "frmPumps.frx":2A32
         textCB          =   "frmPumps.frx":2A4A
         textRB          =   "frmPumps.frx":2A62
         colorBack       =   "frmPumps.frx":2A7A
         colorIntern     =   "frmPumps.frx":2AA4
         colorMO         =   "frmPumps.frx":2ACE
         colorFocus      =   "frmPumps.frx":2AF8
         colorDisabled   =   "frmPumps.frx":2B22
         colorPressed    =   "frmPumps.frx":2B4C
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   11
         Left            =   1230
         TabIndex        =   83
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":2B76
         textLT          =   "frmPumps.frx":2BE6
         textCT          =   "frmPumps.frx":2BFE
         textRT          =   "frmPumps.frx":2C16
         textLM          =   "frmPumps.frx":2C2E
         textRM          =   "frmPumps.frx":2C46
         textLB          =   "frmPumps.frx":2C5E
         textCB          =   "frmPumps.frx":2C76
         textRB          =   "frmPumps.frx":2C8E
         colorBack       =   "frmPumps.frx":2CA6
         colorIntern     =   "frmPumps.frx":2CD0
         colorMO         =   "frmPumps.frx":2CFA
         colorFocus      =   "frmPumps.frx":2D24
         colorDisabled   =   "frmPumps.frx":2D4E
         colorPressed    =   "frmPumps.frx":2D78
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   11
         Left            =   1380
         TabIndex        =   95
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   11
         Left            =   1380
         TabIndex        =   94
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   11
         Left            =   1380
         TabIndex        =   93
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   11
         Left            =   1350
         TabIndex        =   92
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   11
         Left            =   1350
         TabIndex        =   91
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   11
         Left            =   1350
         TabIndex        =   90
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   10
         Left            =   60
         TabIndex        =   89
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   10
         Left            =   60
         TabIndex        =   88
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   10
         Left            =   60
         TabIndex        =   87
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   86
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   85
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   84
         Top             =   840
         Width           =   525
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   5
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   4
      Left            =   2640
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   64
      Top             =   2490
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   4
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No6"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":2DA2
         textLT          =   "frmPumps.frx":2E12
         textCT          =   "frmPumps.frx":2E2A
         textRT          =   "frmPumps.frx":2E42
         textLM          =   "frmPumps.frx":2E5A
         textRM          =   "frmPumps.frx":2E72
         textLB          =   "frmPumps.frx":2E8A
         textCB          =   "frmPumps.frx":2EA2
         textRB          =   "frmPumps.frx":2EBA
         colorBack       =   "frmPumps.frx":2ED2
         colorIntern     =   "frmPumps.frx":2EFC
         colorMO         =   "frmPumps.frx":2F26
         colorFocus      =   "frmPumps.frx":2F50
         colorDisabled   =   "frmPumps.frx":2F7A
         colorPressed    =   "frmPumps.frx":2FA4
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   8
         Left            =   0
         TabIndex        =   66
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":2FCE
         textLT          =   "frmPumps.frx":303E
         textCT          =   "frmPumps.frx":3056
         textRT          =   "frmPumps.frx":306E
         textLM          =   "frmPumps.frx":3086
         textRM          =   "frmPumps.frx":309E
         textLB          =   "frmPumps.frx":30B6
         textCB          =   "frmPumps.frx":30CE
         textRB          =   "frmPumps.frx":30E6
         colorBack       =   "frmPumps.frx":30FE
         colorIntern     =   "frmPumps.frx":3128
         colorMO         =   "frmPumps.frx":3152
         colorFocus      =   "frmPumps.frx":317C
         colorDisabled   =   "frmPumps.frx":31A6
         colorPressed    =   "frmPumps.frx":31D0
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   9
         Left            =   1230
         TabIndex        =   67
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":31FA
         textLT          =   "frmPumps.frx":326A
         textCT          =   "frmPumps.frx":3282
         textRT          =   "frmPumps.frx":329A
         textLM          =   "frmPumps.frx":32B2
         textRM          =   "frmPumps.frx":32CA
         textLB          =   "frmPumps.frx":32E2
         textCB          =   "frmPumps.frx":32FA
         textRB          =   "frmPumps.frx":3312
         colorBack       =   "frmPumps.frx":332A
         colorIntern     =   "frmPumps.frx":3354
         colorMO         =   "frmPumps.frx":337E
         colorFocus      =   "frmPumps.frx":33A8
         colorDisabled   =   "frmPumps.frx":33D2
         colorPressed    =   "frmPumps.frx":33FC
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   9
         Left            =   1380
         TabIndex        =   79
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   9
         Left            =   1380
         TabIndex        =   78
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   9
         Left            =   1380
         TabIndex        =   77
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   9
         Left            =   1350
         TabIndex        =   76
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   9
         Left            =   1350
         TabIndex        =   75
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   9
         Left            =   1350
         TabIndex        =   74
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   8
         Left            =   60
         TabIndex        =   73
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   8
         Left            =   60
         TabIndex        =   72
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   8
         Left            =   60
         TabIndex        =   71
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   70
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   69
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   68
         Top             =   840
         Width           =   525
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   4
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   3
      Left            =   2640
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   48
      Top             =   30
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   3
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No2"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":3426
         textLT          =   "frmPumps.frx":3496
         textCT          =   "frmPumps.frx":34AE
         textRT          =   "frmPumps.frx":34C6
         textLM          =   "frmPumps.frx":34DE
         textRM          =   "frmPumps.frx":34F6
         textLB          =   "frmPumps.frx":350E
         textCB          =   "frmPumps.frx":3526
         textRB          =   "frmPumps.frx":353E
         colorBack       =   "frmPumps.frx":3556
         colorIntern     =   "frmPumps.frx":3580
         colorMO         =   "frmPumps.frx":35AA
         colorFocus      =   "frmPumps.frx":35D4
         colorDisabled   =   "frmPumps.frx":35FE
         colorPressed    =   "frmPumps.frx":3628
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   50
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":3652
         textLT          =   "frmPumps.frx":36C2
         textCT          =   "frmPumps.frx":36DA
         textRT          =   "frmPumps.frx":36F2
         textLM          =   "frmPumps.frx":370A
         textRM          =   "frmPumps.frx":3722
         textLB          =   "frmPumps.frx":373A
         textCB          =   "frmPumps.frx":3752
         textRB          =   "frmPumps.frx":376A
         colorBack       =   "frmPumps.frx":3782
         colorIntern     =   "frmPumps.frx":37AC
         colorMO         =   "frmPumps.frx":37D6
         colorFocus      =   "frmPumps.frx":3800
         colorDisabled   =   "frmPumps.frx":382A
         colorPressed    =   "frmPumps.frx":3854
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   7
         Left            =   1230
         TabIndex        =   51
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":387E
         textLT          =   "frmPumps.frx":38EE
         textCT          =   "frmPumps.frx":3906
         textRT          =   "frmPumps.frx":391E
         textLM          =   "frmPumps.frx":3936
         textRM          =   "frmPumps.frx":394E
         textLB          =   "frmPumps.frx":3966
         textCB          =   "frmPumps.frx":397E
         textRB          =   "frmPumps.frx":3996
         colorBack       =   "frmPumps.frx":39AE
         colorIntern     =   "frmPumps.frx":39D8
         colorMO         =   "frmPumps.frx":3A02
         colorFocus      =   "frmPumps.frx":3A2C
         colorDisabled   =   "frmPumps.frx":3A56
         colorPressed    =   "frmPumps.frx":3A80
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   3
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   63
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   62
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   61
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   7
         Left            =   60
         TabIndex        =   60
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   7
         Left            =   60
         TabIndex        =   59
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   7
         Left            =   60
         TabIndex        =   58
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   6
         Left            =   1350
         TabIndex        =   57
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   6
         Left            =   1350
         TabIndex        =   56
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   6
         Left            =   1350
         TabIndex        =   55
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   6
         Left            =   1380
         TabIndex        =   54
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   6
         Left            =   1380
         TabIndex        =   53
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   6
         Left            =   1380
         TabIndex        =   52
         Top             =   840
         Width           =   525
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   2
      Left            =   30
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   32
      Top             =   4950
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   2
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No9"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":3AAA
         textLT          =   "frmPumps.frx":3B1A
         textCT          =   "frmPumps.frx":3B32
         textRT          =   "frmPumps.frx":3B4A
         textLM          =   "frmPumps.frx":3B62
         textRM          =   "frmPumps.frx":3B7A
         textLB          =   "frmPumps.frx":3B92
         textCB          =   "frmPumps.frx":3BAA
         textRB          =   "frmPumps.frx":3BC2
         colorBack       =   "frmPumps.frx":3BDA
         colorIntern     =   "frmPumps.frx":3C04
         colorMO         =   "frmPumps.frx":3C2E
         colorFocus      =   "frmPumps.frx":3C58
         colorDisabled   =   "frmPumps.frx":3C82
         colorPressed    =   "frmPumps.frx":3CAC
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   34
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":3CD6
         textLT          =   "frmPumps.frx":3D46
         textCT          =   "frmPumps.frx":3D5E
         textRT          =   "frmPumps.frx":3D76
         textLM          =   "frmPumps.frx":3D8E
         textRM          =   "frmPumps.frx":3DA6
         textLB          =   "frmPumps.frx":3DBE
         textCB          =   "frmPumps.frx":3DD6
         textRB          =   "frmPumps.frx":3DEE
         colorBack       =   "frmPumps.frx":3E06
         colorIntern     =   "frmPumps.frx":3E30
         colorMO         =   "frmPumps.frx":3E5A
         colorFocus      =   "frmPumps.frx":3E84
         colorDisabled   =   "frmPumps.frx":3EAE
         colorPressed    =   "frmPumps.frx":3ED8
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   5
         Left            =   1230
         TabIndex        =   35
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":3F02
         textLT          =   "frmPumps.frx":3F72
         textCT          =   "frmPumps.frx":3F8A
         textRT          =   "frmPumps.frx":3FA2
         textLM          =   "frmPumps.frx":3FBA
         textRM          =   "frmPumps.frx":3FD2
         textLB          =   "frmPumps.frx":3FEA
         textCB          =   "frmPumps.frx":4002
         textRB          =   "frmPumps.frx":401A
         colorBack       =   "frmPumps.frx":4032
         colorIntern     =   "frmPumps.frx":405C
         colorMO         =   "frmPumps.frx":4086
         colorFocus      =   "frmPumps.frx":40B0
         colorDisabled   =   "frmPumps.frx":40DA
         colorPressed    =   "frmPumps.frx":4104
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   2
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   47
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   46
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   45
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   5
         Left            =   60
         TabIndex        =   44
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   5
         Left            =   60
         TabIndex        =   43
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   5
         Left            =   60
         TabIndex        =   42
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   4
         Left            =   1350
         TabIndex        =   41
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   4
         Left            =   1350
         TabIndex        =   40
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   4
         Left            =   1350
         TabIndex        =   39
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   4
         Left            =   1380
         TabIndex        =   38
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   4
         Left            =   1380
         TabIndex        =   37
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   4
         Left            =   1380
         TabIndex        =   36
         Top             =   840
         Width           =   525
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   1
      Left            =   30
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   16
      Top             =   2490
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No5"
         Enabled         =   0   'False
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":412E
         textLT          =   "frmPumps.frx":419E
         textCT          =   "frmPumps.frx":41B6
         textRT          =   "frmPumps.frx":41CE
         textLM          =   "frmPumps.frx":41E6
         textRM          =   "frmPumps.frx":41FE
         textLB          =   "frmPumps.frx":4216
         textCB          =   "frmPumps.frx":422E
         textRB          =   "frmPumps.frx":4246
         colorBack       =   "frmPumps.frx":425E
         colorIntern     =   "frmPumps.frx":4288
         colorMO         =   "frmPumps.frx":42B2
         colorFocus      =   "frmPumps.frx":42DC
         colorDisabled   =   "frmPumps.frx":4306
         colorPressed    =   "frmPumps.frx":4330
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   18
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":435A
         textLT          =   "frmPumps.frx":43CA
         textCT          =   "frmPumps.frx":43E2
         textRT          =   "frmPumps.frx":43FA
         textLM          =   "frmPumps.frx":4412
         textRM          =   "frmPumps.frx":442A
         textLB          =   "frmPumps.frx":4442
         textCB          =   "frmPumps.frx":445A
         textRB          =   "frmPumps.frx":4472
         colorBack       =   "frmPumps.frx":448A
         colorIntern     =   "frmPumps.frx":44B4
         colorMO         =   "frmPumps.frx":44DE
         colorFocus      =   "frmPumps.frx":4508
         colorDisabled   =   "frmPumps.frx":4532
         colorPressed    =   "frmPumps.frx":455C
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   3
         Left            =   1230
         TabIndex        =   19
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
         Enabled         =   0   'False
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":4586
         textLT          =   "frmPumps.frx":45F6
         textCT          =   "frmPumps.frx":460E
         textRT          =   "frmPumps.frx":4626
         textLM          =   "frmPumps.frx":463E
         textRM          =   "frmPumps.frx":4656
         textLB          =   "frmPumps.frx":466E
         textCB          =   "frmPumps.frx":4686
         textRB          =   "frmPumps.frx":469E
         colorBack       =   "frmPumps.frx":46B6
         colorIntern     =   "frmPumps.frx":46E0
         colorMO         =   "frmPumps.frx":470A
         colorFocus      =   "frmPumps.frx":4734
         colorDisabled   =   "frmPumps.frx":475E
         colorPressed    =   "frmPumps.frx":4788
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   1
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   31
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   30
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   3
         Left            =   60
         TabIndex        =   28
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   3
         Left            =   60
         TabIndex        =   27
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   3
         Left            =   60
         TabIndex        =   26
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   1350
         TabIndex        =   25
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   1350
         TabIndex        =   24
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   1350
         TabIndex        =   23
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   21
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   20
         Top             =   840
         Width           =   525
      End
   End
   Begin VB.PictureBox picPump 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   0
      Left            =   30
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   30
      Width           =   2535
      Begin BTNENHLib4.BtnEnh cmdPump 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2490
         _Version        =   524298
         _ExtentX        =   4392
         _ExtentY        =   767
         _StockProps     =   66
         Caption         =   "PUMP No1"
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Surface         =   11
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":47B2
         textLT          =   "frmPumps.frx":4822
         textCT          =   "frmPumps.frx":483A
         textRT          =   "frmPumps.frx":4852
         textLM          =   "frmPumps.frx":486A
         textRM          =   "frmPumps.frx":4882
         textLB          =   "frmPumps.frx":489A
         textCB          =   "frmPumps.frx":48B2
         textRB          =   "frmPumps.frx":48CA
         colorBack       =   "frmPumps.frx":48E2
         colorIntern     =   "frmPumps.frx":490C
         colorMO         =   "frmPumps.frx":4936
         colorFocus      =   "frmPumps.frx":4960
         colorDisabled   =   "frmPumps.frx":498A
         colorPressed    =   "frmPumps.frx":49B4
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   420
         Width           =   1230
         _Version        =   524298
         _ExtentX        =   2170
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "UNLEADED"
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":49DE
         textLT          =   "frmPumps.frx":4A4E
         textCT          =   "frmPumps.frx":4A66
         textRT          =   "frmPumps.frx":4A7E
         textLM          =   "frmPumps.frx":4A96
         textRM          =   "frmPumps.frx":4AAE
         textLB          =   "frmPumps.frx":4AC6
         textCB          =   "frmPumps.frx":4ADE
         textRB          =   "frmPumps.frx":4AF6
         colorBack       =   "frmPumps.frx":4B0E
         colorIntern     =   "frmPumps.frx":4B38
         colorMO         =   "frmPumps.frx":4B62
         colorFocus      =   "frmPumps.frx":4B8C
         colorDisabled   =   "frmPumps.frx":4BB6
         colorPressed    =   "frmPumps.frx":4BE0
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin BTNENHLib4.BtnEnh cmdTop 
         Height          =   375
         Index           =   2
         Left            =   1230
         TabIndex        =   3
         Top             =   420
         Width           =   1260
         _Version        =   524298
         _ExtentX        =   2222
         _ExtentY        =   661
         _StockProps     =   66
         Caption         =   "LRP FUEL"
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
         Surface         =   5
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         UserData        =   0.1
         textCaption     =   "frmPumps.frx":4C0A
         textLT          =   "frmPumps.frx":4C7A
         textCT          =   "frmPumps.frx":4C92
         textRT          =   "frmPumps.frx":4CAA
         textLM          =   "frmPumps.frx":4CC2
         textRM          =   "frmPumps.frx":4CDA
         textLB          =   "frmPumps.frx":4CF2
         textCB          =   "frmPumps.frx":4D0A
         textRB          =   "frmPumps.frx":4D22
         colorBack       =   "frmPumps.frx":4D3A
         colorIntern     =   "frmPumps.frx":4D64
         colorMO         =   "frmPumps.frx":4D8E
         colorFocus      =   "frmPumps.frx":4DB8
         colorDisabled   =   "frmPumps.frx":4DE2
         colorPressed    =   "frmPumps.frx":4E0C
         SpotlightOffsetX=   14
         SpotlightOffsetY=   14
         SpotlightResizeWidth=   70
         SpotlightResizeHeight=   20
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   15
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   14
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   12
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   11
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   10
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLitre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblRand 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblcop 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Litres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rands"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   840
         Width           =   525
      End
      Begin MSForms.Image Image1 
         Height          =   1845
         Index           =   0
         Left            =   1200
         Top             =   570
         Width           =   75
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "132;3254"
      End
   End
End
Attribute VB_Name = "frmPumps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTop_Click(index As Integer)
    Unload Me
End Sub
