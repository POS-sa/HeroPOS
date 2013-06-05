VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBank 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Banking Details..."
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAccountName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1770
      TabIndex        =   0
      Top             =   165
      Width           =   3945
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1770
      TabIndex        =   1
      Top             =   525
      Width           =   3945
   End
   Begin VB.TextBox txtBranchName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1770
      TabIndex        =   2
      Top             =   870
      Width           =   3945
   End
   Begin VB.TextBox txtBranchCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1770
      TabIndex        =   3
      Top             =   1215
      Width           =   2115
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1770
      TabIndex        =   4
      Top             =   1560
      Width           =   2115
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   345
      Left            =   4560
      TabIndex        =   6
      ToolTipText     =   " Click to Search.... "
      Top             =   1980
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdOk 
      Height          =   345
      Left            =   3300
      TabIndex        =   5
      ToolTipText     =   " Click to Search.... "
      Top             =   1980
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin MSForms.Image Image7 
      Height          =   285
      Index           =   2
      Left            =   1620
      Top             =   120
      Width           =   4155
      BackColor       =   16777215
      Size            =   "7329;503"
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Account Name:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   12
      Left            =   120
      TabIndex        =   10
      Top             =   885
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Branch Name:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   11
      Left            =   120
      TabIndex        =   9
      Top             =   1230
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Branch Code:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   540
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Bank:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1530
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Account Number:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image7 
      Height          =   285
      Index           =   0
      Left            =   1620
      Top             =   480
      Width           =   4155
      BackColor       =   16777215
      Size            =   "7329;503"
   End
   Begin MSForms.Image Image8 
      Height          =   285
      Left            =   1620
      Top             =   825
      Width           =   4155
      BackColor       =   16777215
      MousePointer    =   1
      Size            =   "7329;503"
   End
   Begin MSForms.Image Image9 
      Height          =   285
      Left            =   1620
      Top             =   1170
      Width           =   2325
      BackColor       =   16777215
      Size            =   "4101;503"
   End
   Begin MSForms.Image Image7 
      Height          =   285
      Index           =   1
      Left            =   1620
      Top             =   1515
      Width           =   2325
      BackColor       =   16777215
      Size            =   "4101;503"
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ActiveUpdateServer "Delete from Bank_Details"
    DoEvents
    ActiveUpdateServer "INSERT INTO [Bank_Details]([Account_Name], [Bank_Name], [Branch_Name], [Branch_Code], [Account_No])" & _
    " VALUES('" & txtAccountName.Text & "','" & txtBankName.Text & "','" & txtBranchName.Text & "','" & txtBranchCode.Text & "','" & txtAccountNo.Text & "')"
    DoEvents
    Unload Me
End Sub
Private Sub Form_Load()
    ActiveReadServer "Select *from Bank_Details"
    If rs.RecordCount > 0 Then
        txtAccountName.Text = rs.Fields("Account_Name")
        txtBankName.Text = rs.Fields("Bank_Name")
        txtBranchName.Text = rs.Fields("Branch_Name")
        txtBranchCode.Text = rs.Fields("Branch_Code")
        txtAccountNo.Text = rs.Fields("Account_No")
    End If
    rs.Close
End Sub

Private Sub txtAccountName_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtAccountName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub

Private Sub txtAccountNo_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub
Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub
Private Sub txtBankName_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub
Private Sub txtBankName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub
Private Sub txtBranchCode_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub
Private Sub txtBranchCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub

Private Sub txtBranchName_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtBranchName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub
