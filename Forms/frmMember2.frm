VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmMember 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1425
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMember2.frx":0000
   ScaleHeight     =   1425
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer errTimer 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   900
      Index           =   0
      Left            =   8040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   2163158
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
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   885
      Left            =   270
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
      Width           =   7695
      _Version        =   524298
      _ExtentX        =   13573
      _ExtentY        =   1561
      _StockProps     =   66
      Caption         =   "Invalid Key Pressed"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
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
      Shape           =   1
      CornerFactor    =   15
      BackColorContainer=   8438015
      SpecialEffect   =   3
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "frmMember2.frx":290A
      textLT          =   "frmMember2.frx":2990
      textCT          =   "frmMember2.frx":29A8
      textRT          =   "frmMember2.frx":29C0
      textLM          =   "frmMember2.frx":29D8
      textRM          =   "frmMember2.frx":29F0
      textLB          =   "frmMember2.frx":2A08
      textCB          =   "frmMember2.frx":2A20
      textRB          =   "frmMember2.frx":2A38
      colorBack       =   "frmMember2.frx":2A50
      colorIntern     =   "frmMember2.frx":2A7A
      colorMO         =   "frmMember2.frx":2AA4
      colorFocus      =   "frmMember2.frx":2ACE
      colorDisabled   =   "frmMember2.frx":2AF8
      colorPressed    =   "frmMember2.frx":2B22
   End
   Begin VB.PictureBox picHoldFocus 
      Height          =   495
      Left            =   8160
      ScaleHeight     =   435
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   390
      Width           =   405
   End
   Begin VB.Label lblMember 
      BackStyle       =   0  'Transparent
      Caption         =   "Please scan the Membership Card."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   450
      TabIndex        =   4
      Top             =   570
      Width           =   6525
   End
   Begin MSForms.TextBox txtNum 
      Height          =   525
      Left            =   510
      TabIndex        =   3
      Top             =   510
      Visible         =   0   'False
      Width           =   4815
      VariousPropertyBits=   746604563
      ForeColor       =   16777215
      Size            =   "8493;926"
      PasswordChar    =   42
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click(Index As Integer)
    KeyRegister = ""
    Unload Me
End Sub
Private Sub cmdErr_Click()
    cmdErr.Caption = ""
    cmdErr.Visible = False
    errTimer.Enabled = False
    picHoldFocus.SetFocus
End Sub
Private Sub errTimer_Timer()
    Select Case cmdErr.BackColor
        Case &HC0C0&          'White
            cmdErr.BackColor = &HFF&
        Case &HFF&             'Yellow
            cmdErr.BackColor = &HC0C0&
    End Select
End Sub

Private Sub Form_Initialize()
Screen.MousePointer = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then
        KeyCode = 0
        Exit Sub
    End If
    If KeyCode > 47 And KeyCode < 58 Then
        lblMember.Visible = False
        txtNum.Visible = True
        If txtNum.Text = "" Then txtNum.Text = Chr(KeyCode)
    End If
    If KeyCode > 95 And KeyCode < 106 Then
        lblMember.Visible = False
        txtNum.Visible = True
        If txtNum.Text = "" Then
            Select Case KeyCode
                Case 96: txtNum.Text = "0"
                Case 97: txtNum.Text = "1"
                Case 98: txtNum.Text = "2"
                Case 99: txtNum.Text = "3"
                Case 100: txtNum.Text = "4"
                Case 101: txtNum.Text = "5"
                Case 101: txtNum.Text = "6"
                Case 103: txtNum.Text = "7"
                Case 104: txtNum.Text = "8"
                Case 104: txtNum.Text = "9"
            End Select
            KeyCode = 0
        End If
    End If
    If KeyCode = 13 Then
        ActiveReadServer "Select * from Debtors where Debt_Type = 4 and Debtor_No ='" & txtNum.Text & "'"
        If rs.RecordCount > 0 Then
        TillData.Creditbalance = rs.Fields("Balance")
        TillData.Creditlimit = rs.Fields("Credit_Limit")
            KeyRegister = txtNum.Text
            Unload Me
        Else
            cmdErr.Caption = "Unknown Membership Card."
            errTimer.Enabled = True
            picHoldFocus.SetFocus
            lblMember.Visible = True
            txtNum.Visible = False
            txtNum.Text = ""
            frmMember.Refresh
            DoEvents
            cmdErr.Visible = True
        End If
        rs.Close
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 32
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtNum_Change()
    If txtNum.Text = "" Then
        lblMember.Visible = True
        txtNum.Visible = False
    End If
End Sub
Private Sub txtNum_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 32
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub
