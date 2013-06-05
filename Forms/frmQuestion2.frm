VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "GIF89.DLL"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmQuestion2 
   BorderStyle     =   0  'None
   Caption         =   "Please Note"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   1440
      Left            =   3210
      OleObjectBlob   =   "frmQuestion2.frx":0000
      TabIndex        =   0
      Top             =   870
      Width           =   1440
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   1170
      Index           =   0
      Left            =   2580
      TabIndex        =   1
      Top             =   2910
      Width           =   2295
      _Version        =   524298
      _ExtentX        =   4048
      _ExtentY        =   2064
      _StockProps     =   66
      Caption         =   "No"
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
      BackColorContainer=   3119822
      ButtonRaiseFactor=   3
      SmoothEdges     =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      FlatPillowFactor=   3
      UserData        =   0.1
      textCaption     =   "frmQuestion2.frx":DA46
      textLT          =   "frmQuestion2.frx":DAAA
      textCT          =   "frmQuestion2.frx":DAC2
      textRT          =   "frmQuestion2.frx":DADA
      textLM          =   "frmQuestion2.frx":DAF2
      textRM          =   "frmQuestion2.frx":DB0A
      textLB          =   "frmQuestion2.frx":DB22
      textCB          =   "frmQuestion2.frx":DB3A
      textRB          =   "frmQuestion2.frx":DB52
      colorBack       =   "frmQuestion2.frx":DB6A
      colorIntern     =   "frmQuestion2.frx":DB94
      colorMO         =   "frmQuestion2.frx":DBBE
      colorFocus      =   "frmQuestion2.frx":DBE8
      colorDisabled   =   "frmQuestion2.frx":DC12
      colorPressed    =   "frmQuestion2.frx":DC3C
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   1170
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   2910
      Width           =   2295
      _Version        =   524298
      _ExtentX        =   4048
      _ExtentY        =   2064
      _StockProps     =   66
      Caption         =   "Yes"
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
      BackColorContainer=   3119822
      ButtonRaiseFactor=   3
      SmoothEdges     =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      FlatPillowFactor=   3
      UserData        =   0.1
      textCaption     =   "frmQuestion2.frx":DC66
      textLT          =   "frmQuestion2.frx":DCCC
      textCT          =   "frmQuestion2.frx":DCE4
      textRT          =   "frmQuestion2.frx":DCFC
      textLM          =   "frmQuestion2.frx":DD14
      textRM          =   "frmQuestion2.frx":DD2C
      textLB          =   "frmQuestion2.frx":DD44
      textCB          =   "frmQuestion2.frx":DD5C
      textRB          =   "frmQuestion2.frx":DD74
      colorBack       =   "frmQuestion2.frx":DD8C
      colorIntern     =   "frmQuestion2.frx":DDB6
      colorMO         =   "frmQuestion2.frx":DDE0
      colorFocus      =   "frmQuestion2.frx":DE0A
      colorDisabled   =   "frmQuestion2.frx":DE34
      colorPressed    =   "frmQuestion2.frx":DE5E
      HollowFrame     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HeroPOS Information Message"
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
   Begin MSForms.Image Image2 
      Height          =   1485
      Left            =   30
      Top             =   2730
      Width           =   4995
      BorderColor     =   16777215
      BackColor       =   3119822
      BorderStyle     =   0
      Size            =   "8811;2619"
   End
   Begin MSForms.Image Image1 
      Height          =   2115
      Left            =   2850
      Top             =   630
      Width           =   2175
      BackColor       =   0
      Size            =   "3836;3731"
   End
   Begin MSForms.Label lblCap 
      Height          =   1905
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   2655
      VariousPropertyBits=   8388627
      Size            =   "4683;3360"
      FontName        =   "Arial Narrow"
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Image Image4 
      Height          =   3615
      Left            =   0
      Top             =   630
      Width           =   5055
      BackColor       =   8638191
      Size            =   "8916;6376"
   End
   Begin MSForms.Image Image3 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   5055
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "8916;1085"
      Picture         =   "frmQuestion2.frx":DE88
   End
End
Attribute VB_Name = "frmQuestion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnEnh1_Click(index As Integer)
    Select Case BtnEnh1(index).Caption
        Case "Yes"
            frmQuestion2.Hide
            Screen.MousePointer = 11
           frmSplash.BtnEnh2.Tag = "Dothecapture"
           
            Exit Sub
       
       
        Case "No"
            Screen.MousePointer = 11
            
            frmSplash.BtnEnh2.Tag = "Dotheclockout"

           
      
            
    End Select
   
    Screen.MousePointer = 1
    Unload Me
End Sub
Private Sub Form_Activate()
    frmQuestion2.lblCap = "Would you like to finalize your Cashup now?"
    Screen.MousePointer = 1
    
            BtnEnh1(1).Caption = "Yes"
            BtnEnh1(0).Caption = "No"
        
End Sub



