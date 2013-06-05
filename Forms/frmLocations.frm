VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLocations 
   Caption         =   "Locations"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picProducts 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9465
      Index           =   2
      Left            =   0
      ScaleHeight     =   9465
      ScaleWidth      =   13875
      TabIndex        =   0
      Top             =   0
      Width           =   13875
      Begin VSFlex8Ctl.VSFlexGrid grdLoc 
         Height          =   3960
         Left            =   0
         TabIndex        =   4
         Top             =   4890
         Width           =   13705
         _cx             =   24174
         _cy             =   6985
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
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmLocations.frx":0000
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
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Touch Panel Links"
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   2760
         TabIndex        =   5
         Top             =   2460
         Width           =   2265
         Begin MSForms.CheckBox chkPanels 
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   8
            Tag             =   "1"
            Top             =   300
            Width           =   1905
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;661"
            Value           =   "0"
            Caption         =   "Retail Panel"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPanels 
            Height          =   375
            Index           =   1
            Left            =   180
            TabIndex        =   7
            Tag             =   "1"
            Top             =   675
            Width           =   1905
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;661"
            Value           =   "0"
            Caption         =   "Restaurant Panel"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPanels 
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   6
            Tag             =   "1"
            Top             =   1050
            Width           =   1905
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;661"
            Value           =   "0"
            Caption         =   "Bar Panel"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.PictureBox picTopFrame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   0
         ScaleHeight     =   585
         ScaleWidth      =   13785
         TabIndex        =   1
         Top             =   4320
         Width           =   13785
         Begin btButtonEx.ButtonEx cmdUp1 
            Height          =   465
            Left            =   13020
            TabIndex        =   2
            Top             =   30
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   820
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
         Begin MSForms.Image Image6 
            Height          =   90
            Left            =   60
            Top             =   420
            Width           =   3015
            BackColor       =   16761024
            Size            =   "5318;159"
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   3
            Left            =   3120
            Top             =   420
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   4
            Left            =   3450
            Top             =   420
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin MSForms.Image Image4 
            Height          =   90
            Index           =   5
            Left            =   3780
            Top             =   420
            Width           =   285
            BackColor       =   16761024
            Size            =   "503;159"
         End
         Begin VB.Label Label5 
            Caption         =   "Location List."
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
            Left            =   150
            TabIndex        =   3
            Top             =   60
            Width           =   3135
         End
         Begin MSForms.Image Image7 
            Height          =   615
            Left            =   0
            Top             =   -30
            Width           =   13705
            BorderStyle     =   0
            SpecialEffect   =   3
            Size            =   "24174;1085"
         End
      End
      Begin MSForms.Image Image5 
         Height          =   345
         Left            =   840
         Top             =   210
         Width           =   3195
         BackColor       =   16777215
         Size            =   "5636;609"
         VariousPropertyBits=   19
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   35
         Left            =   90
         TabIndex        =   15
         Top             =   660
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
      Begin VB.Line Line1 
         Index           =   6
         X1              =   1410
         X2              =   7590
         Y1              =   750
         Y2              =   750
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   36
         Left            =   1260
         TabIndex        =   14
         Top             =   1740
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Location Type:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   38
         Left            =   1260
         TabIndex        =   13
         Top             =   1395
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Location Name:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label4 
         Height          =   225
         Left            =   1260
         TabIndex        =   12
         Top             =   1050
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Location Number:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   39
         Left            =   900
         TabIndex        =   11
         Top             =   240
         Width           =   3105
         ForeColor       =   0
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Location Details"
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
         Height          =   225
         Index           =   22
         Left            =   240
         TabIndex        =   9
         Top             =   2085
         Width           =   2415
         BackColor       =   -2147483643
         Caption         =   "Month End Stock Behaviour:"
         Size            =   "4260;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Image picSeperate1 
         Height          =   105
         Left            =   0
         Top             =   4200
         Width           =   15195
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "26802;185"
      End
      Begin MSForms.TextBox txtLocName 
         Height          =   285
         Left            =   2760
         TabIndex        =   18
         Top             =   1335
         Width           =   6105
         VariousPropertyBits=   746604563
         MaxLength       =   30
         BorderStyle     =   1
         Size            =   "10769;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtLocNum 
         Height          =   285
         Left            =   2760
         TabIndex        =   17
         Top             =   990
         Width           =   2565
         VariousPropertyBits=   746604563
         MaxLength       =   4
         BorderStyle     =   1
         Size            =   "4524;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbLocType 
         Height          =   285
         Left            =   2760
         TabIndex        =   16
         Top             =   1680
         Width           =   2200
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3881;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbStockBehave 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Top             =   2025
         Width           =   2200
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3881;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUp1_Click()
    Select Case cmdUp1.Caption
        Case "5"
            picSeperate1.top = 0
            picTopFrame1.top = picSeperate1.top + picSeperate1.Height + 30
            grdLoc.top = picSeperate1.top + picSeperate1.Height - 50 + picTopFrame1.Height
            grdLoc.Height = picProducts(2).Height - picSeperate1.Height - picTopFrame1.Height - 550
            grdLoc.SetFocus
            cmdUp1.Caption = 6
        Case "6"
            picTopFrame1.top = 4320
            picSeperate1.top = 4200
            grdLoc.Height = 3960
            grdLoc.top = 4890
            cmdUp1.Caption = "5"
            txtLocNum.SetFocus
    End Select
End Sub
Private Sub cmbStockBehave_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If txtLocNum.Text = "" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub
Private Sub chkPanels_Change(Index As Integer)
    Select Case chkPanels(Index).Value
        Case True
            chkPanels(Index).Tag = "1"
        Case Else
            chkPanels(Index).Tag = "0"
    End Select
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If txtLocNum.Text = "" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub


Private Sub cmbLocType_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
    If txtLocNum.Text = "" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub

Private Sub Form_Activate()
    txtLocNum.SetFocus
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Locations"
End Sub

Private Sub Form_Load()
    Dim x As Printer
    RecordCount = 0
    
    grdLoc.TextMatrix(0, 0) = "Location No."
    grdLoc.TextMatrix(0, 1) = "Location Name"
    grdLoc.TextMatrix(0, 2) = "Location Type"
    grdLoc.TextMatrix(0, 3) = "Stock Roll Behavior"
    frmMain.Toolbar1.Buttons(5).Enabled = False
    Load_Locations
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
End Sub

Private Sub txtLocName_Change()
    If txtLocName.Text = "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = False
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = True
    End If
    If txtLocNum.Text = "" Then frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub
Private Sub txtLocName_KeyPress(KeyAscii As MSForms.ReturnInteger)
   
    Select Case KeyAscii
        Case 33 To 47, 58 To 64, 91 To 96, 123 To 127, 162 To 184, 247, 248, 191
            KeyAscii = 0
        
    End Select

End Sub
Private Sub txtLocName_LostFocus()
    On Error Resume Next
    txtLocName.Text = UCase(Left(txtLocName.Text, 1)) & Mid(txtLocName.Text, 2)
On Error GoTo 0
End Sub

Private Sub txtlocnum_Change()
    If txtLocNum.Text = "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = False
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = True
    End If
    
    Select Case grdLoc.FindRow(txtLocNum.Text, , 0, , 1)
        Case -1
        Case Else
            grdLoc.Row = grdLoc.FindRow(txtLocNum.Text, , 0, , 1)
    End Select
End Sub
Private Sub txtlocnum_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtLocNum_LostFocus()
    On Error Resume Next
    If grdLoc.FindRow(txtLocNum.Text, , 0, , 1) = -1 Then
        response = MsgBox("You have typed a number for a location that do not exist. " & Chr$(13) & "Do you want to create a new Location?", vbYesNo, "HeroPOS")
        If response = vbYes Then
            CreateLocation Val(txtLocNum.Text)
        Else
            ActiveReadServer "Select * From Locations where Location_No=" & grdLoc.TextMatrix(grdLoc.Row, 0)
            If rs.RecordCount > 0 Then
                txtLocNum.Text = rs.Fields("Location_No")
                txtLocName.Text = Trim(rs.Fields("Loc_Name"))
                Select Case rs.Fields("Loc_Type")
                    Case 0: cmbLocType.Text = "Sales Location"
                    Case 1: cmbLocType.Text = "Stock Location"
                    Case 2: cmbLocType.Text = "Expence Location"
                    Case 3: cmbLocType.Text = "Outside Location"
                End Select
                Select Case rs.Fields("Stock_Take")
                    Case 0: cmbStockBehave.Text = "Enter Count"
                    Case 1: cmbStockBehave.Text = "Use System Levels"
                    Case 2: cmbStockBehave.Text = "Reset to Zero"
                End Select
            End If
            rs.Close
            frmMain.Toolbar1.Buttons(4).Enabled = False
        End If
    End If
    On Error GoTo 0
End Sub
Private Sub Load_Locations()
    RecordCount = 0
    grdLoc.ColWidth(0) = grdLoc.Width * 0.2
    grdLoc.ColWidth(1) = grdLoc.Width * 0.5
    grdLoc.ColWidth(2) = grdLoc.Width * 0.15
    grdLoc.ColWidth(3) = grdLoc.Width * 0.14
    cmbStockBehave.Clear
    cmbStockBehave.AddItem "Enter Count"
    cmbStockBehave.AddItem "Use System Levels"
    cmbStockBehave.AddItem "Reset to Zero"
    cmbStockBehave.Text = "Enter Count"
    cmbLocType.Clear
    cmbLocType.AddItem "Sales Location"
    cmbLocType.AddItem "Stock Location"
    cmbLocType.AddItem "Expence Location"
    cmbLocType.AddItem "Outside Location"
    cmbLocType.Text = "Sales Location"
    grdLoc.Rows = 1
    ActiveReadServer "Select * From Locations order by Location_No"
    While Not rs.EOF
        grdLoc.Rows = grdLoc.Rows + 1
        With grdLoc
            .TextMatrix(.Rows - 1, 0) = rs.Fields("Location_No")
            .TextMatrix(.Rows - 1, 1) = rs.Fields("Loc_Name") & ""
            Select Case rs.Fields("Stock_Take")
                Case 0: .TextMatrix(.Rows - 1, 3) = "Enter Count"
                Case 1: .TextMatrix(.Rows - 1, 3) = "Use System Levels"
                Case 2: .TextMatrix(.Rows - 1, 3) = "Reset to Zero"
            End Select
            Select Case rs.Fields("Loc_Type")
                Case 0: .TextMatrix(.Rows - 1, 2) = "Sales Location"
                Case 1: .TextMatrix(.Rows - 1, 2) = "Stock Location"
                Case 2: .TextMatrix(.Rows - 1, 2) = "Expence Location"
                Case 3: .TextMatrix(.Rows - 1, 2) = "Outside Location"
            End Select
        End With
        rs.MoveNext
    Wend
    frmMain.stbBar.Panels(3) = "Records = " & rs.RecordCount
    rs.Close
    
    If grdLoc.Rows > 1 Then grdLoc.Row = 1
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            KeyCode = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
            frmSplash.Show
            frmMain.picProdBar.Visible = False
            frmMain.picAccBar.Visible = False
            frmMain.Hide
    End Select
End Sub
Private Sub grdLoc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If grdLoc.Row > 0 Then
        ActiveReadServer "Select * From Locations where Location_No=" & grdLoc.TextMatrix(grdLoc.Row, 0)
            If rs.RecordCount > 0 Then
                txtLocNum.Text = rs.Fields("Location_No")
                txtLocName.Text = Trim(rs.Fields("Loc_Name"))
                Select Case rs.Fields("Loc_Type")
                    Case 0: cmbLocType.Text = "Sales Location"
                    Case 1: cmbLocType.Text = "Stock Location"
                    Case 2: cmbLocType.Text = "Expence Location"
                    Case 3: cmbLocType.Text = "Outside Location"
                End Select
                Select Case rs.Fields("Stock_Take")
                    Case 0: cmbStockBehave.Text = "Enter Count"
                    Case 1: cmbStockBehave.Text = "Use System Levels"
                    Case 2: cmbStockBehave.Text = "Reset to Zero"
                End Select
            End If
            rs.Close
            chkPanels(0).Value = 0
            chkPanels(1).Value = 0
            chkPanels(2).Value = 0
            ActiveReadServer "Select * From Location_Links  where Location_No=" & grdLoc.TextMatrix(grdLoc.Row, 0)
            If rs.RecordCount > 0 Then
                chkPanels(0).Value = rs.Fields("Panel1")
                chkPanels(1).Value = rs.Fields("Panel2")
                chkPanels(2).Value = rs.Fields("Panel3")
            End If
            rs.Close
            frmMain.Toolbar1.Buttons(4).Enabled = False
            frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
End Sub
