VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmProdFind1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Product Search..."
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLine 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1410
      TabIndex        =   0
      Top             =   180
      Width           =   4455
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   795
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1402
      Appearance      =   3
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   795
      Left            =   6000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1402
      Appearance      =   3
      Caption         =   "Find"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.Image Image1 
      Height          =   795
      Left            =   1290
      Top             =   60
      Width           =   4665
      BackColor       =   16777215
      Size            =   "8229;1402"
   End
End
Attribute VB_Name = "frmProdFind1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
    Unload Me
End Sub
Private Sub ButtonEx2_Click()
    ProductFilter(2) = Trim(txtLine.Text)
    Unload Me
End Sub
Private Sub Form_Load()
    txtLine.Text = ""
End Sub
Private Sub txtLine_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtLine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case KeyCode
        Case 13
            KeyCode = 0
            ButtonEx2_Click
        Case 27
            Unload Me
    End Select
End Sub
