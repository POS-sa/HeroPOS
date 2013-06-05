VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmUsers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   210
      ScaleHeight     =   495
      ScaleWidth      =   4875
      TabIndex        =   24
      Top             =   5430
      Width           =   4875
      Begin VB.Label lblUsers 
         Caption         =   "Users List."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   90
         TabIndex        =   25
         Top             =   0
         Width           =   3135
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   2
         Left            =   3720
         Top             =   360
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   1
         Left            =   3390
         Top             =   360
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   0
         Left            =   3060
         Top             =   360
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image3 
         Height          =   90
         Left            =   0
         Top             =   360
         Width           =   3015
         BackColor       =   16761024
         Size            =   "5318;159"
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdUsers 
      Height          =   4260
      Left            =   0
      TabIndex        =   9
      Top             =   5940
      Width           =   13605
      _cx             =   23998
      _cy             =   7514
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   15329975
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmUsers.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin btButtonEx.ButtonEx btnSetting 
      Height          =   315
      Left            =   3990
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Options..."
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
   Begin btButtonEx.ButtonEx cmdUp 
      Height          =   435
      Left            =   13020
      TabIndex        =   23
      Top             =   5460
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   767
      Appearance      =   3
      Caption         =   "5"
      CaptionOffsetX  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   4
      Left            =   1170
      TabIndex        =   28
      Top             =   4785
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "ID Number:"
      Size            =   "2461;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   315
      Left            =   2700
      TabIndex        =   27
      Top             =   4710
      Width           =   2655
      VariousPropertyBits=   746604571
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "4683;556"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblOrderFilter 
      Caption         =   "order by User_No"
      Height          =   525
      Left            =   14190
      TabIndex        =   26
      Top             =   2190
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSForms.Label Label5 
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   5020
      Width           =   1815
      BackColor       =   -2147483643
      Caption         =   "= Required Fields"
      Size            =   "3201;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5010
      Width           =   195
      BackColor       =   -2147483643
      Caption         =   "*"
      Size            =   "344;450"
      FontName        =   "Tahoma"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblUserNumber 
      Height          =   225
      Left            =   2850
      TabIndex        =   0
      Top             =   1170
      Width           =   975
      BackColor       =   -2147483643
      Size            =   "1720;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton optPassExp 
      Height          =   285
      Index           =   1
      Left            =   2700
      TabIndex        =   6
      Top             =   3540
      Width           =   2235
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "3942;503"
      Value           =   "1"
      Caption         =   "Password Never Expires"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton optPassExp 
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   5
      Top             =   3135
      Width           =   1605
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "2831;556"
      Value           =   "0"
      Caption         =   "Password Expires"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   8
      Left            =   1170
      TabIndex        =   20
      Top             =   3930
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "Gender:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtPassword 
      Height          =   315
      Left            =   2700
      TabIndex        =   4
      Top             =   2730
      Width           =   1605
      VariousPropertyBits=   746604571
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "2831;556"
      PasswordChar    =   42
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   7
      Left            =   1170
      TabIndex        =   19
      Top             =   2790
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "* Password:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtLastName 
      Height          =   315
      Left            =   2700
      TabIndex        =   3
      Top             =   2325
      Width           =   2655
      VariousPropertyBits=   746604571
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "4683;556"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   6
      Left            =   1170
      TabIndex        =   18
      Top             =   2400
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "Last Name:"
      Size            =   "2461;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtFirstName 
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
      VariousPropertyBits=   746604571
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "4683;556"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   1170
      TabIndex        =   17
      Top             =   1980
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "First Name:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label4 
      Height          =   885
      Left            =   5820
      TabIndex        =   16
      Top             =   3060
      Width           =   2175
      ForeColor       =   14737632
      BackColor       =   -2147483643
      Caption         =   "Click to add Picture"
      Size            =   "3836;1561"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image Image5 
      Height          =   3195
      Left            =   5460
      Top             =   1920
      Width           =   2835
      BackColor       =   16777215
      Size            =   "5001;5636"
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   15
      Top             =   210
      Width           =   3105
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "User Details"
      Size            =   "5477;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   1170
      TabIndex        =   14
      Top             =   1575
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "* Log on Name:"
      Size            =   "2461;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1590
      X2              =   7020
      Y1              =   900
      Y2              =   900
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   13
      Top             =   720
      Width           =   1785
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "General"
      Size            =   "3149;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   1170
      TabIndex        =   12
      Top             =   1140
      Width           =   1395
      BackColor       =   16777215
      Caption         =   "User Number:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Left            =   1170
      TabIndex        =   11
      Top             =   4380
      Width           =   1455
      BackColor       =   16777215
      Caption         =   "Position:"
      Size            =   "2566;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ComboBox cmbUserType 
      Height          =   315
      Left            =   2700
      TabIndex        =   8
      Top             =   4320
      Width           =   2235
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3942;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbGender 
      Height          =   315
      Left            =   2700
      TabIndex        =   7
      Top             =   3915
      Width           =   1635
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2884;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtLogin 
      Height          =   315
      Left            =   2700
      TabIndex        =   1
      Top             =   1515
      Width           =   4035
      VariousPropertyBits=   746604571
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "7117;556"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image2 
      Height          =   105
      Left            =   0
      Top             =   5310
      Width           =   15195
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "26802;185"
   End
   Begin MSForms.Image picTopbar 
      Height          =   585
      Left            =   0
      Top             =   5400
      Width           =   13605
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "23998;1032"
   End
   Begin MSForms.Image frTop 
      Height          =   345
      Left            =   1380
      Top             =   180
      Width           =   3195
      BackColor       =   16777215
      Size            =   "5636;609"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   315
      Left            =   2700
      Top             =   1110
      Width           =   1215
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "2143;556"
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnSetting_Click()
    frmUserAccess.Show vbModal
    If grdUsers.Col > 1 Then
        grdUsers.Col = 1
        Exit Sub
    End If
    If grdUsers.Col < 5 Then
        grdUsers.Col = 4
        Exit Sub
    End If
End Sub
Private Sub cmbGender_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If Trim(txtLogin.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtPassword.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtLogin.Text) = "Startup" Or Trim(txtLogin.Text) = "New User" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub
Private Sub cmbUserType_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If Trim(txtLogin.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtPassword.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
If Trim(txtLogin.Text) = "Startup" Or Trim(txtLogin.Text) = "New User" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub
Public Sub cmdUp_Click()
    Select Case cmdUp.Caption
        Case "5"
            Image2.top = 0
            picTopbar.top = 90
            grdUsers.top = 630
            Picture1.top = 120
            cmdUp.Caption = "5"
            cmdUp.Caption = 6
            cmdUp.top = 150
            grdUsers.Height = frmUsers.Height - 900
        Case "6"
            Image2.top = 5310
            picTopbar.top = 5400
            grdUsers.top = 5940
            Picture1.top = 5430
            cmdUp.top = 5460
            cmdUp.Caption = "5"
            grdUsers.Height = 4260
    End Select
End Sub
Private Sub Form_Activate()
    If frmUsers.Tag = "1" Then
        frmUsers.Tag = ""
        Exit Sub
    End If
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Users"
    grdUsers.Width = frmUsers.Width - 120
    grdUsers.Height = frmUsers.Height - grdUsers.top - 120
    picTopbar.Width = grdUsers.Width
    Image2.Width = grdUsers.Width
    cmbGender.Clear
    cmbGender.AddItem "Male"
    cmbGender.AddItem "Female"
    cmbGender.Text = "Male"
    cmbUserType.Clear
    cmbUserType.AddItem "Manager"
    cmbUserType.AddItem "Night Manager"
    cmbUserType.AddItem "Reservations Clerk"
    cmbUserType.AddItem "Waiter"
    cmbUserType.AddItem "Barman"
    cmbUserType.AddItem "GRV Clerk"
    cmbUserType.AddItem "Buyer"
    cmbUserType.AddItem "Supervisor"
    cmbUserType.AddItem "Cashier"
    cmbUserType.AddItem "Owner"
    cmbUserType.AddItem "Staff Member"
    cmbUserType.Text = "Manager"
    grdUsers.Rows = 1
    ActiveReadServer "Select User_no, User_Name, First_Name, Last_Name, isnull(Gender,0) as Gender, isnull(User_Type,0) as User_Type From Users order by User_No"
    While Not rs.EOF
        grdUsers.Rows = grdUsers.Rows + 1
        With grdUsers
            .TextMatrix(.Rows - 1, 0) = rs.Fields("User_No")
            .TextMatrix(.Rows - 1, 1) = rs.Fields("User_Name") & ""
            .TextMatrix(.Rows - 1, 2) = rs.Fields("First_Name") & " " & rs.Fields("Last_Name") & ""
            Select Case rs.Fields("Gender")
                Case False: .TextMatrix(.Rows - 1, 3) = "Male"
                Case True: .TextMatrix(.Rows - 1, 3) = "Female"
            End Select
            Select Case rs.Fields("User_Type")
                Case 0: .TextMatrix(.Rows - 1, 4) = "Manager"
                Case 1: .TextMatrix(.Rows - 1, 4) = "Night Manager"
                Case 2: .TextMatrix(.Rows - 1, 4) = "Reservations Clerk"
                Case 3: .TextMatrix(.Rows - 1, 4) = "Waiter"
                Case 4: .TextMatrix(.Rows - 1, 4) = "Barman"
                Case 5: .TextMatrix(.Rows - 1, 4) = "GRV Clerk"
                Case 6: .TextMatrix(.Rows - 1, 4) = "Buyer"
                Case 7: .TextMatrix(.Rows - 1, 4) = "Supervisor"
                Case 8: .TextMatrix(.Rows - 1, 4) = "Cashier"
                Case 9: .TextMatrix(.Rows - 1, 4) = "Owner"
                Case 10: .TextMatrix(.Rows - 1, 4) = "Staff Member"
            End Select
        End With
        rs.MoveNext
    Wend
    frmMain.stbBar.Panels(3) = "Records = " & Val(rs.RecordCount)
    rs.Close
    If grdUsers.Rows > 1 Then grdUsers.Row = 1
    frmMain.Toolbar1.Buttons(2).Enabled = True
    lblOrderFilter.Caption = " order by User_No"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            For i = 0 To frmMain.cmdBar.Count - 1
                frmMain.cmdBar(i).Enabled = True
            Next i
            UserRecord.Password = ""
            KeyCode = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
            Unload frmUsers
            frmSplash.Show
            frmMain.Hide
    End Select
End Sub
Private Sub Form_Load()
    grdUsers.TextMatrix(0, 0) = "No."
    grdUsers.TextMatrix(0, 1) = "Login Name"
    grdUsers.TextMatrix(0, 2) = "User Name"
    grdUsers.TextMatrix(0, 3) = "Gender"
    grdUsers.TextMatrix(0, 4) = "Type"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
End Sub

Private Sub grdUsers_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
     If grdUsers.Row > 0 Then
           ActiveReadServer "Select * From Users where User_No=" & grdUsers.TextMatrix(grdUsers.Row, 0)
            If rs.RecordCount > 0 Then
                btnSetting.Enabled = True
                lblUserNumber.Caption = rs.Fields("User_No")
                txtLogin.Text = Trim(rs.Fields("User_Name"))
                txtFirstName.Text = Trim(rs.Fields("First_Name"))
                txtLastName.Text = Trim(rs.Fields("Last_Name"))
                txtPassword.Text = Trim(rs.Fields("User_Password"))
                Select Case rs.Fields("Gender")
                    Case False: cmbGender.Text = "Male"
                    Case True: cmbGender.Text = "Female"
                    Case Else: cmbGender.Text = "Male"
                End Select
                Select Case rs.Fields("Expires")
                    Case False: optPassExp(0).Value = False: optPassExp(1).Value = True
                    Case True: optPassExp(0).Value = True: optPassExp(1).Value = False
                    Case Else: optPassExp(0).Value = False: optPassExp(1).Value = True
                End Select
                Select Case rs.Fields("User_Type")
                    Case 0: cmbUserType.Text = "Manager"
                    Case 1: cmbUserType.Text = "Night Manager"
                    Case 2: cmbUserType.Text = "Reservations Clerk"
                    Case 3: cmbUserType.Text = "Waiter"
                    Case 4: cmbUserType.Text = "Barman"
                    Case 5: cmbUserType.Text = "GRV Clerk"
                    Case 6: cmbUserType.Text = "Buyer"
                    Case 7: cmbUserType.Text = "Supervisor"
                    Case 8: cmbUserType.Text = "Cashier"
                    Case 9: cmbUserType.Text = "Owner"
                    Case 10: cmbUserType.Text = "Staff Member"
                End Select
            End If
            Load frmUserAccess
            frmUserAccess.lblLoginName = "  " & txtLogin.Text
            frmUserAccess.lblNumber = lblUserNumber.Caption
            Select Case rs.Fields("Ua_reservations")
                Case True: frmUserAccess.auReservation.Value = 1
                Case False: frmUserAccess.auReservation.Value = 0
                Case Else: frmUserAccess.auReservation.Value = 1
            End Select
            
            Select Case rs.Fields("Quotes")
                Case True: frmUserAccess.chkBox(27).Value = 1
                Case False: frmUserAccess.chkBox(27).Value = 0
                Case Else: frmUserAccess.chkBox(27).Value = 1
            End Select
            
            
            Select Case rs.Fields("Ua_Rooms")
                Case True: frmUserAccess.auRooms.Value = 1
                Case False: frmUserAccess.auRooms.Value = 0
                Case Else: frmUserAccess.auRooms.Value = 1
            End Select
            Select Case rs.Fields("Ua_Guests")
                Case True: frmUserAccess.auGuests.Value = 1
                Case False: frmUserAccess.auGuests.Value = 0
                Case Else: frmUserAccess.auGuests.Value = 1
            End Select
            Select Case rs.Fields("Ua_checkin")
                Case True: frmUserAccess.auCheckin.Value = 1
                Case False: frmUserAccess.auCheckin.Value = 0
                Case Else: frmUserAccess.auCheckin.Value = 1
            End Select
            Select Case rs.Fields("Ua_checkout")
                Case True: frmUserAccess.auCheckout.Value = 1
                Case False: frmUserAccess.auCheckout.Value = 0
                Case Else: frmUserAccess.auCheckout.Value = 1
            End Select
            Select Case rs.Fields("Ua_users")
                Case True: frmUserAccess.auUsers.Value = 1
                Case False: frmUserAccess.auUsers.Value = 0
                Case Else: frmUserAccess.auUsers.Value = 1
            End Select
            Select Case rs.Fields("Ua_reports")
                Case True: frmUserAccess.auReports.Value = 1
                Case False: frmUserAccess.auReports.Value = 0
                Case Else: frmUserAccess.auReports.Value = 1
            End Select
            Select Case rs.Fields("Ua_settings")
                Case True: frmUserAccess.auSettings.Value = 1
                Case False: frmUserAccess.auSettings.Value = 0
                Case Else: frmUserAccess.auSettings.Value = 1
            End Select
            Select Case rs.Fields("Ua_Sales")
                Case True: frmUserAccess.auSales.Value = 1
                Case False: frmUserAccess.auSales.Value = 0
                Case Else: frmUserAccess.auSales.Value = 1
            End Select
            Select Case rs.Fields("Ua_Inventory")
                Case True: frmUserAccess.auInventory.Value = 1
                Case False: frmUserAccess.auInventory.Value = 0
                Case Else: frmUserAccess.auInventory.Value = 1
            End Select
            Select Case rs.Fields("Cash_Sales")
                Case True: frmUserAccess.chkBox(0).Value = 1
                Case False: frmUserAccess.chkBox(0).Value = 0
                Case Else: frmUserAccess.chkBox(0).Value = 1
            End Select
            Select Case rs.Fields("Cheque_Sales")
                Case True: frmUserAccess.chkBox(1).Value = 1
                Case False: frmUserAccess.chkBox(1).Value = 0
                Case Else: frmUserAccess.chkBox(1).Value = 1
            End Select
            Select Case rs.Fields("Card_Sales")
                Case True: frmUserAccess.chkBox(2).Value = 1
                Case False: frmUserAccess.chkBox(2).Value = 0
                Case Else: frmUserAccess.chkBox(2).Value = 1
            End Select
            Select Case rs.Fields("Charge_Sales")
                Case True: frmUserAccess.chkBox(3).Value = 1
                Case False: frmUserAccess.chkBox(3).Value = 0
                Case Else: frmUserAccess.chkBox(3).Value = 1
            End Select
            Select Case rs.Fields("Loyalty_Sales")
                Case True: frmUserAccess.chkBox(4).Value = 1
                Case False: frmUserAccess.chkBox(4).Value = 0
                Case Else: frmUserAccess.chkBox(4).Value = 1
            End Select
            Select Case rs.Fields("Item_Corrects")
                Case True: frmUserAccess.chkBox(5).Value = 1
                Case False: frmUserAccess.chkBox(5).Value = 0
                Case Else: frmUserAccess.chkBox(5).Value = 1
            End Select
            Select Case rs.Fields("Voids")
                Case True: frmUserAccess.chkBox(6).Value = 1
                Case False: frmUserAccess.chkBox(6).Value = 0
                Case Else: frmUserAccess.chkBox(6).Value = 1
            End Select
            Select Case rs.Fields("Returns")
                Case True: frmUserAccess.chkBox(7).Value = 1
                Case False: frmUserAccess.chkBox(7).Value = 0
                Case Else: frmUserAccess.chkBox(7).Value = 1
            End Select
            Select Case rs.Fields("Ullages")
                Case True: frmUserAccess.chkBox(24).Value = 1
                Case False: frmUserAccess.chkBox(24).Value = 0
                Case Else: frmUserAccess.chkBox(24).Value = 1
            End Select
            Select Case rs.Fields("Disc_Perc")
                Case True: frmUserAccess.chkBox(8).Value = 1
                Case False: frmUserAccess.chkBox(8).Value = 0
                Case Else: frmUserAccess.chkBox(8).Value = 1
            End Select
            Select Case rs.Fields("Disc_Amt")
                Case True: frmUserAccess.chkBox(9).Value = 1
                Case False: frmUserAccess.chkBox(9).Value = 0
                Case Else: frmUserAccess.chkBox(9).Value = 1
            End Select
            Select Case rs.Fields("Over_Tender")
                Case True: frmUserAccess.chkBox(10).Value = 1
                Case False: frmUserAccess.chkBox(10).Value = 0
                Case Else: frmUserAccess.chkBox(10).Value = 1
            End Select
            Select Case rs.Fields("Payouts")
                Case True: frmUserAccess.chkBox(11).Value = 1
                Case False: frmUserAccess.chkBox(11).Value = 0
                Case Else: frmUserAccess.chkBox(11).Value = 1
            End Select
            Select Case rs.Fields("Pickups")
                Case True: frmUserAccess.chkBox(12).Value = 1
                Case False: frmUserAccess.chkBox(12).Value = 0
                Case Else: frmUserAccess.chkBox(12).Value = 1
            End Select
            Select Case rs.Fields("Loans")
                Case True: frmUserAccess.chkBox(13).Value = 1
                Case False: frmUserAccess.chkBox(13).Value = 0
                Case Else: frmUserAccess.chkBox(13).Value = 1
            End Select
            Select Case rs.Fields("Receive_Acc")
                Case True: frmUserAccess.chkBox(14).Value = 1
                Case False: frmUserAccess.chkBox(14).Value = 0
                Case Else: frmUserAccess.chkBox(14).Value = 1
            End Select
            Select Case rs.Fields("Trans_Store")
                Case True: frmUserAccess.chkBox(15).Value = 1
                Case False: frmUserAccess.chkBox(15).Value = 0
                Case Else: frmUserAccess.chkBox(15).Value = 1
            End Select
            Select Case rs.Fields("Split_Tenders")
                Case True: frmUserAccess.chkBox(16).Value = 1
                Case False: frmUserAccess.chkBox(16).Value = 0
                Case Else: frmUserAccess.chkBox(16).Value = 1
            End Select
            Select Case rs.Fields("Trans_Clear")
                Case True: frmUserAccess.chkBox(17).Value = 1
                Case False: frmUserAccess.chkBox(17).Value = 0
                Case Else: frmUserAccess.chkBox(17).Value = 1
            End Select
            Select Case rs.Fields("Transfer")
                Case True: frmUserAccess.chkBox(18).Value = 1
                Case False: frmUserAccess.chkBox(18).Value = 0
                Case Else: frmUserAccess.chkBox(18).Value = 1
            End Select
            Select Case rs.Fields("Override")
                Case True: frmUserAccess.chkBox(19).Value = 1
                Case False: frmUserAccess.chkBox(19).Value = 0
                Case Else: frmUserAccess.chkBox(19).Value = 1
            End Select
            Select Case rs.Fields("Buffer_Print")
                Case True: frmUserAccess.chkBox(20).Value = 1
                Case False: frmUserAccess.chkBox(20).Value = 0
                Case Else: frmUserAccess.chkBox(20).Value = 1
            End Select
            Select Case rs.Fields("Reprint")
                Case True: frmUserAccess.chkBox(28).Value = 1
                Case False: frmUserAccess.chkBox(28).Value = 0
                Case Else: frmUserAccess.chkBox(28).Value = 0
            End Select
            Select Case rs.Fields("App_Exit")
                Case True: frmUserAccess.chkBox(23).Value = 1
                Case False: frmUserAccess.chkBox(23).Value = 0
                Case Else: frmUserAccess.chkBox(23).Value = 0
            End Select
            Select Case rs.Fields("Search")
                Case True: frmUserAccess.chkBox(21).Value = 1
                Case False: frmUserAccess.chkBox(21).Value = 0
                Case Else: frmUserAccess.chkBox(21).Value = 1
            End Select
            Select Case rs.Fields("Total_Clear")
                Case True: frmUserAccess.chkBox(22).Value = 1
                Case False: frmUserAccess.chkBox(22).Value = 0
                Case Else: frmUserAccess.chkBox(22).Value = 0
            End Select
            Select Case rs.Fields("No_Sales")
                Case True: frmUserAccess.chkBox(25).Value = 1
                Case False: frmUserAccess.chkBox(25).Value = 0
                Case Else: frmUserAccess.chkBox(25).Value = 0
            End Select
            Select Case rs.Fields("Logged_In")
                Case True: frmUserAccess.chkLog.Value = 1
                Case False: frmUserAccess.chkLog.Value = 0
                Case Else: frmUserAccess.chkLog.Value = 1
            End Select
            Select Case rs.Fields("All_Tables")
                Case True: frmUserAccess.chkAllTables.Value = 1
                Case False: frmUserAccess.chkAllTables.Value = 0
                Case Else: frmUserAccess.chkAllTables.Value = 0
            End Select
            Select Case rs.Fields("Draw_Cash")
                Case True: frmUserAccess.chkPanels(0).Value = 1
                Case False: frmUserAccess.chkPanels(0).Value = 0
                Case Else: frmUserAccess.chkPanels(0).Value = 0
            End Select
            Select Case rs.Fields("Draw_Card")
                Case True: frmUserAccess.chkPanels(1).Value = 1
                Case False: frmUserAccess.chkPanels(1).Value = 0
                Case Else: frmUserAccess.chkPanels(1).Value = 0
            End Select
            Select Case rs.Fields("Draw_Cheque")
                Case True: frmUserAccess.chkPanels(2).Value = 1
                Case False: frmUserAccess.chkPanels(2).Value = 0
                Case Else: frmUserAccess.chkPanels(2).Value = 0
            End Select
            Select Case rs.Fields("Draw_Charge")
                Case True: frmUserAccess.chkPanels(3).Value = 1
                Case False: frmUserAccess.chkPanels(3).Value = 0
                Case Else: frmUserAccess.chkPanels(3).Value = 0
            End Select
            Select Case rs.Fields("Draw_Loyalty")
                Case True: frmUserAccess.chkPanels(4).Value = 1
                Case False: frmUserAccess.chkPanels(4).Value = 0
                Case Else: frmUserAccess.chkPanels(4).Value = 0
            End Select
            Select Case Val(rs.Fields("Com_Calc") & "")
                Case 1: frmUserAccess.chkComms.Value = 1
                Case 0: frmUserAccess.chkComms.Value = 0
                Case Else: frmUserAccess.chkComms.Value = 0
            End Select
            Select Case Val(rs.Fields("Bar_Cash") & "")
                Case 1: frmUserAccess.chkBarCash.Value = 1
                Case 0: frmUserAccess.chkBarCash.Value = 0
                Case Else: frmUserAccess.chkBarCash.Value = 0
            End Select
            
            Select Case Val(rs.Fields("Service_Charge") & "")
                Case 1: frmUserAccess.chkBox(26).Value = 1
                Case 0: frmUserAccess.chkBox(26).Value = 0
                Case Else: frmUserAccess.chkBox(26).Value = 0
            End Select
            'Kotie 22-03-2013 13:15
            frmUserAccess.chkBox(29).Value = rs.Fields("Owner_Transfer")

            
            
            frmUserAccess.txtWage = Format(Val(rs.Fields("Wage") & ""), "0.00")
            frmUserAccess.txtComm1 = Val(rs.Fields("Comm1") & "")
            frmUserAccess.txtComm2 = Val(rs.Fields("Comm2") & "")
            rs.Close
            Service_Charge = 0
             frmMain.Toolbar1.Buttons(4).Enabled = False
             frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
End Sub

Private Sub grdUsers_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If Trim(txtLogin.Text) = "New User" Then
        ActiveUpdateServer "Delete from Users where User_No = " & lblUserNumber
        DeleteRow = grdUsers.Row
        txtLogin.Text = ""
        grdUsers.RemoveItem (DeleteRow)
        frmMain.Toolbar1.Buttons(2).Enabled = True
     End If
     On Error GoTo 0
End Sub

Private Sub grdUsers_BeforeSort(ByVal Col As Long, Order As Integer)
    Select Case Trim(grdUsers.TextMatrix(0, Col))
        Case "No."
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by User_No"
                Case 2
                    lblOrderFilter.Caption = " Order by User_No Desc"
            End Select
        Case "Login Name"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [Login_Name]"
                Case 2
                    lblOrderFilter.Caption = " Order by [Login_Name] Desc"
            End Select
        Case "User Name"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [User_Name]"
                Case 2
                    lblOrderFilter.Caption = " Order by [User_Name] Desc"
            End Select
        Case "Gender"
             Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [Gender]"
                Case 2
                    lblOrderFilter.Caption = " Order by [Gender] Desc"
            End Select
        Case "Type"
            Select Case Order
                Case 1
                    lblOrderFilter.Caption = " Order by [Type]"
                Case 2
                    lblOrderFilter.Caption = " Order by [Type] Desc"
            End Select
    End Select
End Sub

Private Sub optPassExp_Change(Index As Integer)
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If Trim(txtLogin.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtPassword.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtLogin.Text) = "Startup" Or Trim(txtLogin.Text) = "New User" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub txtFirstName_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If Trim(txtLogin.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtPassword.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtLogin.Text) = "Startup" Or Trim(txtLogin.Text) = "New User" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As MSForms.ReturnInteger)
 Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtFirstName_LostFocus()
On Error Resume Next
txtFirstName.Text = UCase(Left(txtFirstName.Text, 1)) & Mid(txtFirstName.Text, 2)
On Error GoTo 0
End Sub

Private Sub txtLastName_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If Trim(txtLogin.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtPassword.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtLogin.Text) = "Startup" Or Trim(txtLogin.Text) = "New User" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub
Private Sub txtLastName_KeyPress(KeyAscii As MSForms.ReturnInteger)
 Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtLastName_LostFocus()
    On Error Resume Next
    txtLastName.Text = UCase(Left(txtLastName.Text, 1)) & Mid(txtLastName.Text, 2)
    DoEvents
    On Error GoTo 0
End Sub

Private Sub txtLogin_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If Trim(txtLogin.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtPassword.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtLogin.Text) = "Startup" Or Trim(txtLogin.Text) = "New User" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub
Private Sub txtLogin_KeyPress(KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtLogin_LostFocus()
    On Error Resume Next
    txtLogin.Text = UCase(Left(txtLogin.Text, 1)) & Mid(txtLogin.Text, 2)
    On Error GoTo 0
End Sub
Private Sub txtPassword_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If Trim(txtLogin.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtPassword.Text = "") Then frmMain.Toolbar1.Buttons(4).Enabled = False
    If Trim(txtLogin.Text) = "Startup" Or Trim(txtLogin.Text) = "New User" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8, 27, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

