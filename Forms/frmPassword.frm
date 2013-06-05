VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "BUTTONEX.OCX"
Begin VB.Form frmPassword 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4770
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1350
      Top             =   2160
   End
   Begin btButtonEx.ButtonEx cmdInvalid 
      Height          =   4065
      Left            =   90
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   7170
      Appearance      =   3
      BackColor       =   255
      Caption         =   "Invalid User !!!"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   10
      Left            =   1320
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   5
      Left            =   2550
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   2
      Left            =   2550
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   4
      Left            =   1320
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   1
      Left            =   1320
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   6
      Left            =   90
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   3
      Left            =   90
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   0
      Left            =   90
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   11
      Left            =   2550
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   8
      Left            =   2550
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   7
      Left            =   1320
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1230
   End
   Begin VB.PictureBox cmdInput 
      Height          =   1020
      Index           =   9
      Left            =   90
      ScaleHeight     =   960
      ScaleWidth      =   1170
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1230
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   4230
      Width           =   3705
      VariousPropertyBits=   746604571
      MaxLength       =   6
      BorderStyle     =   1
      Size            =   "6535;820"
      PasswordChar    =   42
      SpecialEffect   =   0
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInput_Click(Index As Integer)
    TextBox1.SetFocus
    Select Case cmdInput(Index).Caption
        Case "Ok"
            ActiveReadServer "select * from Users where User_password ='" & TextBox1.Text & "'"
            Select Case rs.RecordCount
                Case 1
                    frmSales.lblUsername.Caption = rs.Fields("User_Name")
                    Me.Hide
                Case Else
                    cmdInvalid.Visible = True
                    Timer1.Enabled = True
            End Select
            rs.Close
            
        Case "CL"
            TextBox1.Text = ""
        Case Else
            SendKeys cmdInput(Index).Caption
    End Select
End Sub
Private Sub cmdInvalid_Click()
    cmdInvalid.Visible = False
    Timer1.Enabled = False
    TextBox1.Text = ""
    TextBox1.SetFocus
End Sub

Private Sub TextBox1_Change()
If TextBox1.Text = "" Then
    cmdInput(11).Enabled = False
Else
    cmdInput(11).Enabled = True
End If
End Sub

Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            Me.Hide
        Case 8
            TextBox1.Text = ""
    End Select
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57
        Case Else
            KeyAscii = 0
     End Select
End Sub

Private Sub Timer1_Timer()
    Select Case cmdInvalid.BackColor
        Case &HFF&
            cmdInvalid.BackColor = &HC0C0C0
        Case &HC0C0C0
            cmdInvalid.BackColor = &HFF&
    End Select
End Sub
