VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmKeyBoard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   13080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmKeyBoard.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer errTimer 
      Interval        =   500
      Left            =   30
      Top             =   30
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   10
      Left            =   10200
      TabIndex        =   1
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   19
      Left            =   11430
      TabIndex        =   2
      Top             =   2880
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   12
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "Q"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   13
      Left            =   1590
      TabIndex        =   4
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "W"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   14
      Left            =   2820
      TabIndex        =   5
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   15
      Left            =   4050
      TabIndex        =   6
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   16
      Left            =   5280
      TabIndex        =   7
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "T"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   17
      Left            =   6510
      TabIndex        =   8
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "Y"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   26
      Left            =   7740
      TabIndex        =   9
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   27
      Left            =   8970
      TabIndex        =   10
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "I"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   20
      Left            =   1230
      TabIndex        =   11
      Top             =   4200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   21
      Left            =   2460
      TabIndex        =   12
      Top             =   4200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   22
      Left            =   3690
      TabIndex        =   13
      Top             =   4200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   23
      Left            =   4920
      TabIndex        =   14
      Top             =   4200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "F"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   24
      Left            =   6150
      TabIndex        =   15
      Top             =   4200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "G"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   25
      Left            =   7380
      TabIndex        =   16
      Top             =   4200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "H"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   44
      Left            =   8610
      TabIndex        =   17
      Top             =   4200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "J"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   28
      Left            =   9840
      TabIndex        =   18
      Top             =   4200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "K"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   29
      Left            =   11070
      TabIndex        =   19
      Top             =   4200
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1275
      Index           =   40
      Left            =   11040
      TabIndex        =   20
      Top             =   6780
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2249
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   31
      Left            =   1590
      TabIndex        =   21
      Top             =   5490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "Z"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   33
      Left            =   2820
      TabIndex        =   22
      Top             =   5490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   35
      Left            =   4050
      TabIndex        =   23
      Top             =   5490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   36
      Left            =   5280
      TabIndex        =   24
      Top             =   5490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   37
      Left            =   6510
      TabIndex        =   25
      Top             =   5490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   38
      Left            =   7740
      TabIndex        =   26
      Top             =   5490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   39
      Left            =   8970
      TabIndex        =   27
      Top             =   5490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   "M"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   2595
      Index           =   41
      Left            =   10200
      TabIndex        =   28
      Top             =   4200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4577
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "Enter"
      CaptionOffsetY  =   45
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   32
      Left            =   360
      TabIndex        =   29
      Top             =   5490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "Home"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1275
      Index           =   34
      Left            =   3630
      TabIndex        =   30
      Top             =   6780
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   2249
      Appearance      =   3
      BackColor       =   10736617
      BorderColor     =   8421504
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1275
      Index           =   43
      Left            =   360
      TabIndex        =   31
      Top             =   6780
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2249
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "CL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1305
      Index           =   30
      Left            =   360
      TabIndex        =   32
      Top             =   4200
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   2302
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "DEL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1080
      Index           =   11
      Left            =   11760
      TabIndex        =   33
      Top             =   345
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   255
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
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   0
      Left            =   360
      TabIndex        =   34
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   1
      Left            =   1590
      TabIndex        =   35
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   2
      Left            =   2820
      TabIndex        =   36
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   3
      Left            =   4050
      TabIndex        =   37
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   4
      Left            =   5280
      TabIndex        =   38
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   5
      Left            =   6510
      TabIndex        =   39
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   6
      Left            =   7740
      TabIndex        =   40
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   7
      Left            =   8970
      TabIndex        =   41
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   8
      Left            =   10200
      TabIndex        =   42
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1335
      Index           =   9
      Left            =   11430
      TabIndex        =   43
      Top             =   1560
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2355
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   1095
      Left            =   330
      TabIndex        =   44
      Top             =   330
      Visible         =   0   'False
      Width           =   11295
      _Version        =   524298
      _ExtentX        =   19923
      _ExtentY        =   1931
      _StockProps     =   66
      Caption         =   "Invalid Key Pressed"
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
      CornerFactor    =   15
      BackColorContainer=   12632256
      SpecialEffect   =   3
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "frmKeyBoard.frx":74F2
      textLT          =   "frmKeyBoard.frx":7578
      textCT          =   "frmKeyBoard.frx":7590
      textRT          =   "frmKeyBoard.frx":75A8
      textLM          =   "frmKeyBoard.frx":75C0
      textRM          =   "frmKeyBoard.frx":75D8
      textLB          =   "frmKeyBoard.frx":75F0
      textCB          =   "frmKeyBoard.frx":7608
      textRB          =   "frmKeyBoard.frx":7620
      colorBack       =   "frmKeyBoard.frx":7638
      colorIntern     =   "frmKeyBoard.frx":7662
      colorMO         =   "frmKeyBoard.frx":768C
      colorFocus      =   "frmKeyBoard.frx":76B6
      colorDisabled   =   "frmKeyBoard.frx":76E0
      colorPressed    =   "frmKeyBoard.frx":770A
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1275
      Index           =   18
      Left            =   2160
      TabIndex        =   45
      Top             =   6780
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   2249
      Appearance      =   3
      BackColor       =   6208225
      BorderColor     =   8421504
      Caption         =   "*"
      CaptionOffsetY  =   6
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
   Begin MSForms.TextBox txtReg 
      Height          =   735
      Left            =   450
      TabIndex        =   0
      Top             =   600
      Width           =   9795
      VariousPropertyBits=   746604563
      ForeColor       =   16777215
      Size            =   "17277;1296"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   480
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image newBack 
      Height          =   1365
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1365
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "2408;2408"
      Picture         =   "frmKeyBoard.frx":7734
   End
End
Attribute VB_Name = "frmKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdErr_Click()
    errTimer.Enabled = False
    cmdErr.Visible = False
    cmdErr.Caption = ""
    cmdErr.BackColor = &HB0&
End Sub
Private Sub cmdKey_Click(Index As Integer)
      On Error Resume Next
      txtReg.SetFocus
      If Index = 11 Then
            If frmKeyBoard.Tag = "Tabs" Then
                  frmBar.lblKeyRegister.Caption = ""
            End If
            Me.Hide
            On Error GoTo 0
            Exit Sub
      End If
      If cmdErr.Visible = True Then
            If cmdKey(Index).Caption <> "CL" Then
                On Error GoTo 0
                Exit Sub
            End If
      End If
      If Index = 34 Then
          txtReg.Text = txtReg.Text + "  "
      End If
      Select Case cmdKey(Index).Caption
            Case "*"
                SendKeys cmdKey(Index).Caption
            Case "DEL"
                  txtReg.Text = ""
            Case "Home"
                  SendKeys "{HOME}"
            Case "Enter"
                  If frmKeyBoard.Tag = "Message" Then
                        frmSales1.lblPayReason = Mid(Trim(txtReg.Text), 1, 49)
                        Me.Hide
                        On Error GoTo 0
                        Exit Sub
                  End If
                  If frmKeyBoard.Tag = "Payout" Then
                        frmSales.lblPayReason = Trim(txtReg.Text)
                        Me.Hide
                        On Error GoTo 0
                        Exit Sub
                  End If
                  If Trim(txtReg.Text) = "" And frmKeyBoard.Tag = "Tabs" Then
                        cmdErr.Caption = "You have to Enter a Tab Name to Continue"
                        errTimer.Enabled = True
                        cmdErr.Visible = True
                        On Error GoTo 0
                        Exit Sub
                  Else
                       ActiveReadServer "Select * from Tab_Listing where Tab_Name = '" & txtReg.Text & "'"
                       If rs.RecordCount > 0 Then
                            cmdErr.Caption = "Tab already exits. Please change to Continue"
                            errTimer.Enabled = True
                            cmdErr.Visible = True
                            txtReg.Text = ""
                            rs.Close
                            On Error GoTo 0
                            Exit Sub
                       End If
                       rs.Close
                  End If
                  If Trim(txtReg.Text) = "" And frmKeyBoard.Tag = "Tables" Then
                        cmdErr.Caption = "You have to Enter a Table Name to Continue"
                        errTimer.Enabled = True
                        cmdErr.Visible = True
                        On Error GoTo 0
                        Exit Sub
                  Else
                       ActiveReadServer "Select * from Table_Listing where Table_Name = '" & txtReg.Text & "'"
                       If rs.RecordCount > 0 Then
                            cmdErr.Caption = "Table name already exits. Please change to Continue"
                            errTimer.Enabled = True
                            cmdErr.Visible = True
                            txtReg.Text = ""
                            rs.Close
                            On Error GoTo 0
                            Exit Sub
                       End If
                       rs.Close
                       TillData.Table_Name = txtReg.Text
                       Me.Hide
                  End If
                  If frmKeyBoard.Tag = "" Then
                        If frmRestRes.cmdKey(7).Caption = "Name?" And Trim(txtReg.Text) <> "" Then
                              frmRestRes.lblFor = Trim(txtReg.Text)
                        End If
                        If frmRestRes.picPatrons.Visible = True Then frmRestRes.lblMany.Caption = Trim(txtReg.Text)
                        Me.Hide
                        On Error GoTo 0
                        Exit Sub
                  End If
                  If frmKeyBoard.Tag = "Tabs" Then
                        frmBar.lblKeyRegister.Caption = Trim(txtReg.Text)
                        Me.Hide
                        On Error GoTo 0
                        Exit Sub
                  End If
                  If frmKeyBoard.Tag = "frmSales" Then
                        If Left(txtReg.Text, 1) = "*" Then
                         With frmSales
                            .adoData.ConnectionString = cnnMain.ConnectionString
                            .adoData.CursorLocation = adUseServer
                            .adoData.CursorType = adOpenStatic
                            .adoData.LockType = adLockReadOnly
                            .adoData.RecordSource = "Select Product_Code,Description,Department,SOH,Landed_Cost,Tax_Rate,Selling_Price from Product_List where Sales_Item=1 and Description like '%" & Replace(txtReg.Text, "*", "") & "%' order by Description"
                            .adoData.Refresh
                            .grdFind.Col = 1
                            .grdFind.Row = 1
                        End With
                        End If
                        Me.Hide
                        On Error GoTo 0
                        Exit Sub
                  End If
            Case "CL"
                  If errTimer.Enabled = True Then
                        errTimer.Enabled = False
                        cmdErr.Visible = False
                        cmdErr.Caption = ""
                        cmdErr.BackColor = &HB0&
                        On Error GoTo 0
                        Exit Sub
                  End If
                  SendKeys "{BKSP}"
            Case "0" To "9"
                  SendKeys cmdKey(Index).Caption
            Case Else
                  If frmKeyBoard.Tag = "" Then
                    If frmRestRes.picPatrons.Visible = False Then SendKeys cmdKey(Index).Caption
                  Else
                    SendKeys cmdKey(Index).Caption
                  End If
      End Select
      On Error GoTo 0
End Sub
Private Sub errTimer_Timer()
    Select Case cmdErr.BackColor
        Case &HB0&            'White
            cmdErr.BackColor = &H40ADB0
        Case &H40ADB0       'Yellow
            cmdErr.BackColor = &HB0&
    End Select
End Sub
Private Sub Form_Activate()
      On Error Resume Next
      Screen.MousePointer = 0
      If frmSplash.Height < 10000 And newBack.Visible = False Then
        DoEvents
        newBack.Width = Me.Width
        newBack.Height = Me.Height
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            If Me.Controls(i).Name <> "newBack" Then
                Me.Controls(i).Width = Me.Controls(i).Width * 0.798
                Me.Controls(i).Left = Me.Controls(i).Left * 0.79
                Me.Controls(i).Height = Me.Controls(i).Height * 0.792
                Me.Controls(i).top = Me.Controls(i).top * 0.78
                Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
                Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
            End If
        Next i
      End If
      txtReg.Text = ""
      txtReg.SetFocus
      DoEvents
      On Error GoTo 0
End Sub

Private Sub Form_Load()
    If frmSplash.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        DoEvents
        Me.Width = Me.Width * 0.782
        Me.Height = Me.Height * 0.782
        On Error GoTo 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmKeyBoard.Tag = ""
End Sub
Private Sub txtReg_Change()
    If frmKeyBoard.Tag = "frmSales" Then
        If Left(txtReg.Text, 1) = "*" Then
            Exit Sub
        End If
        If frmSales.grdFind.FindRow(txtReg.Text, 0, 1, 0, 0) = -1 Then
            frmSales.grdFind.Row = 1
            txtReg.Text = ""
        Else
            frmSales.grdFind.Row = frmSales.grdFind.FindRow(txtReg.Text, 0, 1, 0, 0)
        End If
    End If
End Sub

Private Sub txtReg_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdKey_Click 41
    End If
    If KeyCode = 27 Then
        cmdKey_Click 11
    End If
End Sub
