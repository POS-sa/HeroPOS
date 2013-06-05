VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "GIF89.DLL"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmServerErr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HeroPOS Server Error Message"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5340
   Icon            =   "frmServerErr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   5340
   StartUpPosition =   1  'CenterOwner
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   1440
      Left            =   3480
      OleObjectBlob   =   "frmServerErr.frx":000C
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   1080
      Left            =   3300
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
      _Version        =   524298
      _ExtentX        =   3201
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
      CornerFactor    =   30
      Surface         =   7
      ButtonRaiseFactor=   3
      SmoothEdges     =   1
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      FlatPillowFactor=   3
      UserData        =   0.1
      textCaption     =   "frmServerErr.frx":DA52
      textLT          =   "frmServerErr.frx":DAB6
      textCT          =   "frmServerErr.frx":DACE
      textRT          =   "frmServerErr.frx":DAE6
      textLM          =   "frmServerErr.frx":DAFE
      textRM          =   "frmServerErr.frx":DB16
      textLB          =   "frmServerErr.frx":DB2E
      textCB          =   "frmServerErr.frx":DB46
      textRB          =   "frmServerErr.frx":DB5E
      colorBack       =   "frmServerErr.frx":DB76
      colorIntern     =   "frmServerErr.frx":DBA0
      colorMO         =   "frmServerErr.frx":DBCA
      colorFocus      =   "frmServerErr.frx":DBF4
      colorDisabled   =   "frmServerErr.frx":DC1E
      colorPressed    =   "frmServerErr.frx":DC48
      HollowFrame     =   -1  'True
   End
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   315
      Left            =   3180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7350
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "&Test Connection"
      Enabled         =   0   'False
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
   Begin MSForms.ComboBox cmbDatabase 
      Height          =   315
      Left            =   1830
      TabIndex        =   10
      Top             =   6900
      Width           =   2775
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4895;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtServer 
      Height          =   315
      Left            =   1830
      TabIndex        =   16
      Top             =   3810
      Width           =   2655
      VariousPropertyBits=   746604571
      Size            =   "4683;556"
      BorderColor     =   16761024
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton opbAuth 
      Height          =   315
      Index           =   0
      Left            =   690
      TabIndex        =   15
      Top             =   4740
      Width           =   2505
      BackColor       =   16777215
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "4419;556"
      Value           =   "0"
      Caption         =   "Use Windows authentication"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton opbAuth 
      Height          =   315
      Index           =   1
      Left            =   690
      TabIndex        =   14
      Top             =   5070
      Width           =   2505
      BackColor       =   16777215
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "4419;556"
      Value           =   "0"
      Caption         =   "Use Server authentication"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Left            =   570
      TabIndex        =   13
      Top             =   5610
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "Login Name:"
      Size            =   "2461;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label3 
      Height          =   315
      Left            =   510
      TabIndex        =   12
      Top             =   5970
      Width           =   1455
      BackColor       =   16777215
      Caption         =   "Password:"
      Size            =   "2566;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label4 
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   11
      Top             =   6930
      Width           =   1605
      BackColor       =   16777215
      Caption         =   "Database:"
      Size            =   "2831;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   315
      Index           =   1
      Left            =   540
      TabIndex        =   9
      Top             =   6480
      Width           =   885
      BackColor       =   16777215
      Caption         =   "Options"
      Size            =   "1561;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   315
      Left            =   1050
      TabIndex        =   8
      Top             =   3840
      Width           =   735
      BackColor       =   16777215
      Caption         =   "Server:"
      Size            =   "1296;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label4 
      Height          =   315
      Index           =   2
      Left            =   300
      TabIndex        =   7
      Top             =   4440
      Width           =   885
      BackColor       =   16777215
      Caption         =   "Connection"
      Size            =   "1561;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtLogin 
      Height          =   315
      Left            =   2010
      TabIndex        =   6
      Top             =   5550
      Width           =   2655
      VariousPropertyBits=   746604571
      Size            =   "4683;556"
      BorderColor     =   16761024
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPassword 
      Height          =   315
      Left            =   2010
      TabIndex        =   5
      Top             =   5910
      Width           =   2655
      VariousPropertyBits=   746604571
      Size            =   "4683;556"
      PasswordChar    =   42
      BorderColor     =   16761024
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line2 
      X1              =   1410
      X2              =   4900
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   4900
      Y1              =   6540
      Y2              =   6540
   End
   Begin MSForms.Label lblCap 
      Height          =   2475
      Left            =   240
      TabIndex        =   3
      Top             =   390
      Width           =   2655
      VariousPropertyBits=   8388627
      Size            =   "4683;4366"
      FontName        =   "Arial Narrow"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Image Image2 
      Height          =   3225
      Left            =   60
      Top             =   60
      Width           =   2985
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "5265;5689"
   End
   Begin MSForms.Image Image1 
      Height          =   4875
      Index           =   0
      Left            =   60
      Top             =   3330
      Width           =   5205
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "9181;8599"
   End
   Begin MSForms.Image Image1 
      Height          =   3225
      Index           =   1
      Left            =   3120
      Top             =   60
      Width           =   2175
      BackColor       =   0
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "3836;5689"
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
      TabIndex        =   2
      Top             =   3630
      Width           =   4785
   End
End
Attribute VB_Name = "frmServerErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnEnh1_Click()
    Server.SQL_Name = txtServer.Text
    Server.SQL_Database = cmbDatabase.Text
    Server.SQL_Password = txtPassword.Text
    Server.SQL_User = txtLogin.Text
    SaveSetting Trim(gblApp_Name), "Server", "Server", txtServer.Text
    SaveSetting Trim(gblApp_Name), "Server", "SQL_User", txtLogin.Text
    SaveSetting Trim(gblApp_Name), "Server", "SQL_Password", txtPassword.Text
    SaveSetting Trim(gblApp_Name), "Server", "SQL_Database", cmbDatabase.Text
    MsgBox "You will have to Restart the application for these settings to take Effect", vbCritical, "HeroPOS Message"
    End
End Sub
Private Sub cmbDatabase_Change()
    Select Case Trim(cmbDatabase.Text)
        Case ""
            cmdForms.Enabled = False
        Case Else
            cmdForms.Enabled = True
    End Select
End Sub
Private Sub cmbDatabase_DropButtonClick()
    Static Dropped
    If Dropped = False Then
        Dropped = True
        ActiveReadServer5 "Exec sp_databases"
        While Not rs5.EOF
            cmbDatabase.AddItem rs5.Fields("Database_Name")
            rs5.MoveNext
        Wend
        rs5.Close
    End If
End Sub
Private Sub cmbDatabase_GotFocus()
    cmbDatabase.SelStart = 0
    cmbDatabase.SelLength = Len(cmbDatabase.Text)
End Sub
Private Sub cmbDatabase_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub cmdForms_Click()
    Screen.MousePointer = 11
    Set cnnMain = New ADODB.Connection
    cnnMain.ConnectionString = "driver={SQL Server};server=" & Trim(txtServer.Text) & ";uid=" & Trim(txtLogin.Text) & ";pwd=" & Trim(txtPassword.Text) & ";database=" & Trim(cmbDatabase.Text)
    cnnMain.ConnectionTimeout = 15
    cnnMain.Open
    If cnnMain.State = 1 Then
        MsgBox "Connection to Server Established...", vbOKOnly, "Server Connection"
    Else
        MsgBox "Connection to Server Failed...", vbOKOnly, "Server Connection"
    End If
    Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
    txtServer.Text = Trim(Server.SQL_Name)
    txtServer.Tag = Trim(Server.SQL_Name)
    txtLogin.Text = Trim(Server.SQL_User)
    txtPassword.Text = Trim(Server.SQL_Password)
    If Trim(Server.SQL_Database) <> "" Then
        cmbDatabase.AddItem Trim(Server.SQL_Database)
        cmbDatabase.Text = Trim(Server.SQL_Database)
        cmbDatabase.Tag = Trim(Server.SQL_Database)
    End If
End Sub
Private Sub txtLogin_GotFocus()
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub
Private Sub txtLogin_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtServer_Change()
    If Trim(txtServer.Text) = "" Then
        cmdForms.Enabled = False
    Else
        cmdForms.Enabled = True
    End If
End Sub
Private Sub txtServer_GotFocus()
    txtServer.SelStart = 0
    txtServer.SelLength = Len(txtServer.Text)
End Sub
Private Sub txtServer_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 32
            KeyAscii = 95
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub
