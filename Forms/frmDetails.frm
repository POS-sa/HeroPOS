VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmDetails 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmLogo 
      Left            =   30
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Logo"
      Filter          =   "*.jpg,*.bmp"
   End
   Begin MSComCtl2.DTPicker tmPicker 
      Height          =   330
      Index           =   0
      Left            =   2520
      TabIndex        =   18
      Top             =   5970
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
      _Version        =   393216
      Format          =   66060290
      CurrentDate     =   38761
   End
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   1365
      Left            =   2670
      TabIndex        =   2
      Top             =   2160
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   2408
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmDetails.frx":0000
   End
   Begin btButtonEx.ButtonEx btnSetting 
      Height          =   315
      Left            =   7740
      TabIndex        =   7
      Top             =   1740
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Settings..."
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
   Begin btButtonEx.ButtonEx cmdRegions 
      Height          =   315
      Left            =   510
      TabIndex        =   8
      Top             =   3660
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Region..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmbTaxRates 
      Height          =   315
      Left            =   480
      TabIndex        =   9
      Top             =   5430
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Tax Rates..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker tmPicker 
      Height          =   345
      Index           =   1
      Left            =   4050
      TabIndex        =   19
      Top             =   5970
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   66060290
      CurrentDate     =   38761
   End
   Begin btButtonEx.ButtonEx cmdSlip 
      Height          =   315
      Left            =   7740
      TabIndex        =   21
      Top             =   2550
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Slip Details..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdNotice 
      Height          =   315
      Left            =   7740
      TabIndex        =   22
      Top             =   2970
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Daily Notice..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdHappy 
      Height          =   315
      Left            =   7740
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Happy Hour..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdVoid 
      Height          =   315
      Left            =   7740
      TabIndex        =   24
      Top             =   4800
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Override Reasons..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtPicker 
      Height          =   330
      Left            =   2520
      TabIndex        =   25
      Top             =   6360
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy MMMM dd"
      Format          =   66060291
      CurrentDate     =   38776
   End
   Begin btButtonEx.ButtonEx cmdBank 
      Height          =   315
      Left            =   480
      TabIndex        =   27
      Top             =   5010
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Banking Details..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdSpecials 
      Height          =   315
      Left            =   7740
      TabIndex        =   28
      Top             =   3390
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Specials..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdWorkStations 
      Height          =   315
      Left            =   7740
      TabIndex        =   29
      Top             =   2130
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Enabled         =   0   'False
      Caption         =   "Workstations..."
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
   Begin btButtonEx.ButtonEx cmdPublic 
      Height          =   315
      Left            =   7740
      TabIndex        =   30
      Top             =   4290
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Public Holidays..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.Label Label4 
      Height          =   885
      Left            =   10830
      TabIndex        =   31
      Top             =   3000
      Width           =   2175
      ForeColor       =   14737632
      BackColor       =   -2147483643
      Caption         =   "Click to add Logo"
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
      Height          =   3405
      Left            =   10230
      Top             =   1740
      Width           =   3375
      BorderColor     =   8421504
      BackColor       =   16777215
      SizeMode        =   1
      Size            =   "5953;6006"
   End
   Begin VB.Line Line2 
      X1              =   7200
      X2              =   7200
      Y1              =   1770
      Y2              =   9060
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   510
      TabIndex        =   26
      Top             =   6420
      Width           =   1785
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Financial Year End Date:"
      Size            =   "3149;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblTo 
      Height          =   255
      Index           =   6
      Left            =   3810
      TabIndex        =   20
      Top             =   6030
      Width           =   225
      BackColor       =   16777215
      Caption         =   "to"
      Size            =   "397;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   510
      TabIndex        =   17
      Top             =   6030
      Width           =   1185
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Business Hours:"
      Size            =   "2090;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image3 
      Height          =   1425
      Left            =   2520
      Top             =   2130
      Width           =   3315
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "5847;2514"
   End
   Begin MSForms.Label lblTax 
      Height          =   225
      Left            =   11730
      TabIndex        =   6
      Top             =   5340
      Width           =   1005
      BackColor       =   16777215
      Caption         =   "1"
      Size            =   "1773;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPrefix 
      Height          =   315
      Left            =   10230
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
      VariousPropertyBits=   746604571
      MaxLength       =   4
      BorderStyle     =   1
      Size            =   "2143;556"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtName 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   1725
      Width           =   4035
      VariousPropertyBits=   746604571
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "7117;556"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtNo 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   4500
      Width           =   1245
      VariousPropertyBits=   746604575
      MaxLength       =   4
      BorderStyle     =   1
      Size            =   "2196;556"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbRegion 
      Height          =   345
      Left            =   2520
      TabIndex        =   3
      Top             =   3645
      Width           =   2685
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4736;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbType 
      Height          =   345
      Left            =   2520
      TabIndex        =   4
      Top             =   4050
      Width           =   2685
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4736;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   7740
      TabIndex        =   16
      Top             =   5340
      Width           =   1815
      BackColor       =   16777215
      Caption         =   "In House Barcode Prefix:"
      Size            =   "3201;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   315
      Left            =   510
      TabIndex        =   15
      Top             =   4110
      Width           =   1125
      BackColor       =   16777215
      Caption         =   "Company Type:"
      Size            =   "1984;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   3
      Left            =   510
      TabIndex        =   14
      Top             =   4590
      Width           =   1185
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Branch Number:"
      Size            =   "2090;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   13
      Top             =   930
      Width           =   855
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "General"
      Size            =   "1508;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1410
      X2              =   12060
      Y1              =   1110
      Y2              =   1110
   End
   Begin MSForms.Label Label3 
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Top             =   2160
      Width           =   615
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Address:"
      Size            =   "1085;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   1770
      Width           =   1185
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Company Name:"
      Size            =   "2090;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Index           =   1
      Left            =   660
      TabIndex        =   10
      Top             =   330
      Width           =   4725
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Company Details"
      Size            =   "8334;556"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image frmTop 
      Height          =   525
      Left            =   510
      Top             =   240
      Width           =   5025
      BackColor       =   16777215
      Size            =   "8864;926"
   End
   Begin MSForms.Image Image2 
      Height          =   315
      Left            =   11610
      Top             =   5280
      Width           =   1215
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "2143;556"
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSetting_Click()
    If UserRecord.Settings = False Then
        MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
        Exit Sub
    End If
    frmSettings.Show vbModal
End Sub

Private Sub cmdBank_Click()
    If UserRecord.Settings = False Then
        MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
        Exit Sub
    End If
    frmBank.Show vbModal
End Sub
Private Sub cmdHappy_Click()
    frmHappy.Show
End Sub
Private Sub cmdNotice_Click()
    frmNotice.Show vbModal
End Sub

Private Sub cmdPublic_Click()
    frmHolidays.Show vbModal
End Sub
Private Sub cmdRegions_Click()
    If UserRecord.Settings = False Then
        MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
        Exit Sub
    End If
    frmRegions.Show vbModal
End Sub
Private Sub cmbTaxRates_Click()
    If UserRecord.Settings = False Then
        MsgBox "           USER ACCESS DENIED", vbApplicationModal, "HeroPOS User Information"
        Exit Sub
    End If
    frmTaxRates.Show vbModal
End Sub

Private Sub cmdSlip_Click()
    frmSlipDetails.Show vbModal
End Sub

Private Sub cmdSpecials_Click()
    frmSpecials.Show vbModal
End Sub
Private Sub cmdVoid_Click()
    frmReasons.Show vbModal
End Sub



Private Sub Form_Activate()
    On Error Resume Next
    DoEvents
    cmbType.Clear
    cmbType.AddItem "Guest House"
    cmbType.AddItem "Boutique Hotel"
    cmbType.AddItem "Lodge"
    cmbType.AddItem "Hotel"
    cmbType.AddItem "Restaurant"
    cmbType.AddItem "Guest House and Restaurant"
    cmbType.AddItem "Club"
    cmbType.AddItem "Supermarket"
    cmbType.AddItem "Butchery"
    cmbType.AddItem "Warehouse"
    cmbType.AddItem "Filling Station"
    cmbType.AddItem "Retail Store"
    tmPicker(0).Value = "08:00:00 AM"
    tmPicker(1).Value = "06:00:00 PM"
    If Trim(Server.SQL_Name) <> "" Then
        ActiveReadServer "Select * from Branch_Details"
        If rs.RecordCount > 0 Then
            dtPicker.Value = rs.Fields("Fin_Year")
            txtNo.Text = rs.Fields("Branch_No")
            txtName.Text = rs.Fields("Branch_Name")
            txtAddress.Text = rs.Fields("Address") & ""
            region = rs.Fields("Region")
            Select Case rs.Fields("Branch_Type")
                Case 1: cmbType.Text = "Guest House": cmdHappy.Visible = True
                Case 2: cmbType.Text = "Boutique Hotel": cmdHappy.Visible = True
                Case 3: cmbType.Text = "Lodge": cmdHappy.Visible = True
                Case 4: cmbType.Text = "Hotel": cmdHappy.Visible = True
                Case 5: cmbType.Text = "Restaurant": cmdHappy.Visible = True
                Case 6: cmbType.Text = "Guest House and Restaurant": cmdHappy.Visible = True
                Case 7: cmbType.Text = "Club": cmdHappy.Visible = True
                Case 8: cmbType.Text = "Supermarket"
                Case 9: cmbType.Text = "Butchery"
                Case 10: cmbType.Text = "Warehouse"
                Case 11: cmbType.Text = "Filling Station"
                Case 12: cmbType.Text = "Retail Store"
                Case Else: cmbType.Text = "Guest House": cmdHappy.Visible = True
            End Select
            lblTax.Caption = rs.Fields("Tax_Rate") & ""
            txtPrefix.Text = Trim(rs.Fields("Prefix") & "")
            CodePrefix = Trim(rs.Fields("Prefix") & "")
            If IsNull(rs.Fields("Time_Start")) Then
                tmPicker(0).Value = "08:00:00 AM"
            Else
                tmPicker(0).Value = rs.Fields("Time_Start")
            End If
            Time_Start = tmPicker(0).Value
            If IsNull(rs.Fields("Time_Stop")) Then
                tmPicker(1).Value = "05:00:00 PM"
            Else
                tmPicker(1).Value = rs.Fields("Time_Stop")
            End If
             Time_Stop = tmPicker(1).Value
        End If
        rs.Close
        cmbRegion.Clear
        ActiveReadServer "Select * from Regions order by Region_No"
        While Not rs.EOF
            cmbRegion.AddItem rs.Fields("Region_No") & " - " & rs.Fields("Region_Name")
            If region = rs.Fields("Region_No") Then
                cmbRegion.Text = rs.Fields("Region_No") & " - " & rs.Fields("Region_Name")
            End If
            rs.MoveNext
        Wend
        rs.Close
    End If
    frmMain.Toolbar1.Buttons(2).Enabled = False
    If UserRecord.Settings = False Then
        frmMain.Toolbar1.Buttons(4).Enabled = False
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = True
    End If
    On Error GoTo 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            UserRecord.Password = ""
            KeyCode = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
            frmSplash.Show
            frmMain.Hide
    End Select
End Sub
Private Sub Form_Load()
    On Error Resume Next
    For Each Form In Forms
        Select Case Form.Name
            Case "frmRooms": frmRooms.Hide
            Case "frmRes": frmRes.Hide
            Case "frmGuests": frmGuests.Hide
            Case "frmCheck": frmCheck.Hide
            Case "frmUsers": frmUsers.Hide
            Case "frmReports": frmReports.Hide
        End Select
    Next Form
    If Logo_File = "" Then Label4.Visible = True Else Label4.Visible = False
    Set Image5.Picture = LoadPicture(Logo_File)
    If Trim(Server.SQL_Name) <> "" Then
        ActiveReadServer "Select count(*) as RecCount from Tax_Rates"
        If rs.Fields("RecCount") = 0 Then
            ActiveUpdateServer "INSERT INTO [Tax_Rates]([Tax_Type], [Tax_Rate],[Description])" & _
            "VALUES(1,14,'Vat')"
        End If
        rs.Close
        
        ActiveReadServer "Select count(*) as RecCount from Tax_Rates"
        If rs.Fields("RecCount") = 0 Then
            ActiveUpdateServer "INSERT INTO [Branch_Details]([Branch_No], [Branch_Name], [Address], [Region], [Branch_Type], [Prefix], [Tax_Rate])" & _
            "VALUES(1,'New Branch','',1,1,'','1')"
        End If
        rs.Close
        
        ActiveReadServer "Select count(*) as RecCount from Regions"
        If rs.Fields("RecCount") = 0 Then
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (1,'Eastern Cape')"
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (2,'Free State')"
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (3,'Gauteng')"
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (4,'KwaZulu Natal')"
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (5,'Limpopo')"
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (6,'Mpumalanga')"
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (7,'North West')"
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (8,'Northern Cape')"
            ActiveUpdateServer "INSERT INTO [Regions]([Region_No], [Region_Name]) VALUES (9,'Western Cape')"
        End If
        rs.Close
    End If
    DoEvents
    On Error GoTo 0
End Sub

Private Sub Image5_Click()
    Load_Logo
End Sub



Private Sub Label4_Click()
     Load_Logo
End Sub
Private Sub Load_Logo()
    On Error Resume Next
    cmLogo.Action = 1
    cmLogo.InitDir = App.Path & "\Images"
    Label4.Visible = False
    Set Image5.Picture = LoadPicture("")
    DoEvents
    Logo_File = cmLogo.FileName
    Set Image5.Picture = LoadPicture(cmLogo.FileName)
    ActiveUpdateServer "Update Branch_Details set Logo_File = '" & cmLogo.FileName & "'"
    frmDetails.Refresh
    On Error GoTo 0
End Sub
Private Sub txtAddress_Change()
On Error Resume Next
    Rows = 0
    For i = 1 To Len(txtAddress.Text)
        If Asc(Mid(txtAddress.Text, i, 1)) = 13 Then
            Rows = Rows + 1
            If Rows = 11 Then
                cmbRegion.SetFocus
                Exit For
            End If
        End If
    Next i
On Error GoTo 0
End Sub
'Private Sub txtAddress_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 39
'            KeyAscii = Asc("`")
'    End Select
'End Sub
Private Sub txtAddress_LostFocus()
    newstring = txtAddress.Text
    For i = 1 To Len(txtAddress.Text)
        If i = 1 Then
            If Asc(Mid(txtAddress.Text, 1, 1)) > 96 And Asc(Mid(txtAddress.Text, 1, 1)) < 123 Then
                Mid(newstring, i, 1) = UCase(Mid(txtAddress.Text, 1, 1))
            End If
        End If
        If Asc(Mid(txtAddress.Text, i, 1)) = 13 Then
            If i + 1 = Len(txtAddress.Text) Then
            Else
                Mid(newstring, i + 2, 1) = UCase(Mid(txtAddress.Text, i + 2, 1))
            End If
        End If
    Next i
    txtAddress.Text = Trim(newstring)
End Sub
'Private Sub txtName_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    Select Case KeyAscii
'        Case 39
'            KeyAscii = Asc("`")
'    End Select
'End Sub
Private Sub txtName_LostFocus()
    On Error Resume Next
    txtName.Text = UCase(Left(txtName.Text, 1)) & Mid(txtName.Text, 2)
    On Error GoTo 0
End Sub
Private Sub txtNo_Change()
    If txtNo.Text = "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = False
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = True
    End If
End Sub
Private Sub txtNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57

        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtNo_LostFocus()
    On Error Resume Next
    txtNo.Text = UCase(Left(txtNo.Text, 1)) & Mid(txtNo.Text, 2)
    On Error GoTo 0
End Sub
Private Sub txtPrefix_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
'        Case 39
'            KeyAscii = Asc("`")
        Case 32
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub
