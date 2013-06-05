VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmbarcode 
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1950
      ScaleHeight     =   465
      ScaleWidth      =   1725
      TabIndex        =   4
      Top             =   3300
      Width           =   1755
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   1170
      Index           =   0
      Left            =   3330
      TabIndex        =   0
      Top             =   1860
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
      textCaption     =   "frmbarcode.frx":0000
      textLT          =   "frmbarcode.frx":0064
      textCT          =   "frmbarcode.frx":007C
      textRT          =   "frmbarcode.frx":0094
      textLM          =   "frmbarcode.frx":00AC
      textRM          =   "frmbarcode.frx":00C4
      textLB          =   "frmbarcode.frx":00DC
      textCB          =   "frmbarcode.frx":00F4
      textRB          =   "frmbarcode.frx":010C
      colorBack       =   "frmbarcode.frx":0124
      colorIntern     =   "frmbarcode.frx":014E
      colorMO         =   "frmbarcode.frx":0178
      colorFocus      =   "frmbarcode.frx":01A2
      colorDisabled   =   "frmbarcode.frx":01CC
      colorPressed    =   "frmbarcode.frx":01F6
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   1170
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   1800
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
      textCaption     =   "frmbarcode.frx":0220
      textLT          =   "frmbarcode.frx":0286
      textCT          =   "frmbarcode.frx":029E
      textRT          =   "frmbarcode.frx":02B6
      textLM          =   "frmbarcode.frx":02CE
      textRM          =   "frmbarcode.frx":02E6
      textLB          =   "frmbarcode.frx":02FE
      textCB          =   "frmbarcode.frx":0316
      textRB          =   "frmbarcode.frx":032E
      colorBack       =   "frmbarcode.frx":0346
      colorIntern     =   "frmbarcode.frx":0370
      colorMO         =   "frmbarcode.frx":039A
      colorFocus      =   "frmbarcode.frx":03C4
      colorDisabled   =   "frmbarcode.frx":03EE
      colorPressed    =   "frmbarcode.frx":0418
      HollowFrame     =   -1  'True
   End
   Begin VB.Label Label2 
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
      Left            =   1350
      TabIndex        =   3
      Top             =   180
      Width           =   3345
   End
   Begin MSForms.Image Image3 
      Height          =   615
      Left            =   420
      Top             =   60
      Width           =   5055
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "8916;1085"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Create and update barcode label images?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   420
      TabIndex        =   2
      Top             =   690
      Width           =   5175
   End
   Begin MSForms.Image Image2 
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   6105
      BorderColor     =   0
      BackColor       =   3119822
      Size            =   "10769;5530"
   End
End
Attribute VB_Name = "frmbarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnEnh1_Click(Index As Integer)
If Index = 0 Then Unload frmbarcode
If Index = 1 Then
On Error GoTo 0
Screen.MousePointer = 11
Dim fso As New FileSystemObject
'**fso.DeleteFolder ("c:\Bogus")
If fso.FolderExists(App.Path & "\Barcodes\") = True Then
Kill App.Path & "\Barcodes\" & "*._"
Else
fso.CreateFolder (App.Path & "\Barcodes\")
End If
Barcode1.Visible = True
If rs.State = 1 Then rs.Close
ActiveReadServer " SELECT * from Products "
Dim rsproductsfield  As String
Dim rsproductssave As String
While Not rs.EOF
    Barcode1.Caption = rs.Fields("Product_Code")
        
        Picture1.Height = Barcode1.Height
        Picture1.Width = Barcode1.Width
        Barcode1.PrinterScaleMode = frmbarcode.ScaleMode
        Barcode1.PrinterWidth = Barcode1.Width
        Barcode1.PrinterHeight = Barcode1.Height
        Barcode1.PrinterTop = 0
        Barcode1.PrinterLeft = 0
        Barcode1.PrinterHDC = Picture1.hDC
        Picture1.Refresh
        SavePicture Picture1.Image, (App.Path & "\Barcodes\" & rs.Fields("Product_Code") & "._")
        rsproductssave = (App.Path & "\Barcodes\" & rs.Fields("Product_Code") & "._")
        rsproductsfield = rs.Fields("Product_Code")
        ActiveUpdateServer "Insert Into  Barclinks (Product_Code, BarCode_Links) values(" & rsproductsfield & ", '" & rsproductssave & " ' ) "
       DoEvents
        rs.MoveNext
    Wend
    Barcode1.Visible = False

    End If
    Screen.MousePointer = 1
    Unload frmbarcode
    Exit Sub
Errors:
    Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
Barcode1.Visible = False
End Sub
