VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "GIF89.DLL"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmError 
   Appearance      =   0  'Flat
   BackColor       =   &H00B3E7F0&
   BorderStyle     =   0  'None
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   1440
      Left            =   3480
      OleObjectBlob   =   "frmError.frx":000C
      TabIndex        =   2
      Top             =   840
      Width           =   1440
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   1080
      Left            =   3240
      TabIndex        =   1
      Top             =   2580
      Width           =   1935
      _Version        =   524298
      _ExtentX        =   3413
      _ExtentY        =   1905
      _StockProps     =   66
      Caption         =   "Ok"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21
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
      CornerFactor    =   100
      Surface         =   7
      BackColorContainer=   4210752
      ButtonRaiseFactor=   3
      SmoothEdges     =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      FlatPillowFactor=   3
      UserData        =   0.1
      textCaption     =   "frmError.frx":DA52
      textLT          =   "frmError.frx":DAB6
      textCT          =   "frmError.frx":DACE
      textRT          =   "frmError.frx":DAE6
      textLM          =   "frmError.frx":DAFE
      textRM          =   "frmError.frx":DB16
      textLB          =   "frmError.frx":DB2E
      textCB          =   "frmError.frx":DB46
      textRB          =   "frmError.frx":DB5E
      colorBack       =   "frmError.frx":DB76
      colorIntern     =   "frmError.frx":DBA0
      colorMO         =   "frmError.frx":DBCA
      colorFocus      =   "frmError.frx":DBF4
      colorDisabled   =   "frmError.frx":DC1E
      colorPressed    =   "frmError.frx":DC48
      HollowFrame     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HeroPOS Error Message"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   210
      TabIndex        =   4
      Top             =   150
      Width           =   4005
   End
   Begin MSForms.Image Image3 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   5325
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "9393;1085"
      Picture         =   "frmError.frx":DC72
   End
   Begin MSForms.Label lblCap 
      Height          =   2475
      Left            =   270
      TabIndex        =   0
      Top             =   930
      Width           =   2655
      VariousPropertyBits=   8388627
      Size            =   "4683;4366"
      FontName        =   "Arial Narrow"
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Image Image1 
      Height          =   3315
      Left            =   3120
      Top             =   480
      Width           =   2175
      BackColor       =   0
      Size            =   "3836;5847"
   End
   Begin MSForms.Image Image4 
      Height          =   3255
      Left            =   0
      Top             =   570
      Width           =   5325
      BorderColor     =   32896
      BackColor       =   11790320
      Size            =   "9393;5741"
   End
   Begin MSForms.Image Image2 
      Height          =   1605
      Left            =   0
      Top             =   3930
      Width           =   5295
      Size            =   "9340;2831"
   End
   Begin VB.Label lblError 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   270
      TabIndex        =   3
      Top             =   4110
      Width           =   4785
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnEnh1_Click()
    ImPrinting = False
    frmSplash.Tag = "Not Now"
    Unload frmError
End Sub

Private Sub Form_Initialize()
If Screen.MousePointer <> 1 Then Screen.MousePointer = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    lblError.Caption = ""
End Sub
