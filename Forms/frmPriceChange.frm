VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPriceChange 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Price Change..."
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLandCost 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   2055
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   1020
      Width           =   1725
   End
   Begin VB.TextBox txtMarkup 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2055
      TabIndex        =   3
      Text            =   "0"
      Top             =   1380
      Width           =   1725
   End
   Begin VB.TextBox txtGross 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2055
      TabIndex        =   4
      Text            =   "0"
      Top             =   1755
      Width           =   1725
   End
   Begin VB.TextBox txtSellExcl 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2055
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   2115
      Width           =   1725
   End
   Begin VB.TextBox txtSellIncl 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2055
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   2805
      Width           =   1725
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   5970
      TabIndex        =   2
      ToolTipText     =   " Click to Search.... "
      Top             =   7110
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Height          =   375
      Left            =   4710
      TabIndex        =   1
      Top             =   7110
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save"
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
   Begin btButtonEx.ButtonEx cmdLabels 
      Height          =   285
      Left            =   3840
      TabIndex        =   20
      Top             =   90
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Print Labels"
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
   Begin VSFlex8Ctl.VSFlexGrid grdGrid 
      Height          =   3930
      Left            =   60
      TabIndex        =   21
      Top             =   3120
      Width           =   7140
      _cx             =   12594
      _cy             =   6932
      Appearance      =   0
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
      BackColorSel    =   16706532
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPriceChange.frx":0000
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
      Editable        =   2
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
   Begin btButtonEx.ButtonEx cmdMore 
      Height          =   375
      Left            =   60
      TabIndex        =   22
      Top             =   7110
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Find Similar Products..."
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
   Begin MSForms.Label Label1 
      Height          =   225
      Left            =   450
      TabIndex        =   19
      Top             =   2790
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Selling Price (incl):"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   4470
      TabIndex        =   18
      Top             =   1050
      Width           =   2595
   End
   Begin VB.Label lblAve 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4470
      TabIndex        =   17
      Top             =   2490
      Width           =   2595
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   0
      Left            =   1950
      Top             =   1320
      Width           =   2325
      BackColor       =   16777215
      Size            =   "4101;529"
   End
   Begin MSForms.ComboBox cmbTax 
      Height          =   285
      Left            =   1950
      TabIndex        =   6
      Tag             =   "Up"
      Top             =   2400
      Width           =   2325
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4101;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   5
      Left            =   450
      TabIndex        =   16
      Top             =   2430
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Tax:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   4
      Left            =   450
      TabIndex        =   15
      Top             =   2085
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Selling Price (excl):"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   3
      Left            =   450
      TabIndex        =   14
      Top             =   1740
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Gross Profit%:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   2
      Left            =   450
      TabIndex        =   13
      Top             =   1395
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Markup%"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCost 
      Height          =   225
      Left            =   450
      TabIndex        =   12
      Top             =   1050
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Landed Cost:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   1
      Left            =   1950
      Top             =   1680
      Width           =   2325
      BackColor       =   16777215
      Size            =   "4101;529"
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   2
      Left            =   1950
      Top             =   2040
      Width           =   2325
      BackColor       =   16777215
      Size            =   "4101;529"
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   3
      Left            =   1950
      Top             =   960
      Width           =   2325
      BackColor       =   16777215
      Size            =   "4101;529"
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   4
      Left            =   1950
      Top             =   2745
      Width           =   2325
      BackColor       =   16777215
      Size            =   "4101;529"
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7380
      Y1              =   840
      Y2              =   840
   End
   Begin MSForms.TextBox txtProductCode 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   90
      Width           =   2325
      VariousPropertyBits=   748701717
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "4101;503"
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtDescription 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   435
      Width           =   5745
      VariousPropertyBits=   746604565
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "10134;503"
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   6
      Left            =   -30
      TabIndex        =   8
      Top             =   150
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Product Code:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   12
      Left            =   -30
      TabIndex        =   7
      Top             =   495
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Description:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image2 
      Height          =   645
      Left            =   4350
      Top             =   2400
      Width           =   2835
      BackColor       =   16777215
      Size            =   "5001;1138"
   End
   Begin MSForms.Image Image1 
      Height          =   1375
      Left            =   4350
      Top             =   960
      Width           =   2835
      BackColor       =   16777215
      Size            =   "5001;2425"
   End
End
Attribute VB_Name = "frmPriceChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTax_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbTax_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtSellExcl.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
End Sub

Private Sub cmdLabels_Click()
    Load frmLabel
    frmLabel.Tag = "PriceChange"
    DoEvents
    frmLabel.Show vbModal
End Sub

Private Sub cmdMore_Click()
    frmSim.Show vbModal
    grdGrid.SetFocus
End Sub
Private Sub Form_Load()
    grdGrid.Cols = 4
    grdGrid.ColWidth(0) = grdGrid.Width * 0.2
    grdGrid.ColWidth(1) = grdGrid.Width * 0.45
    grdGrid.ColWidth(2) = grdGrid.Width * 0.2
    grdGrid.ColWidth(3) = grdGrid.Width * 0.1
    grdGrid.ColAlignment(0) = flexAlignLeftCenter
    grdGrid.ColAlignment(1) = flexAlignLeftCenter
    grdGrid.ColAlignment(2) = flexAlignRightCenter
    grdGrid.ColAlignment(3) = flexAlignCenterCenter
    grdGrid.TextMatrix(0, 0) = "Product Code"
    grdGrid.TextMatrix(0, 1) = "Description"
    grdGrid.TextMatrix(0, 2) = "Selling Price"
    grdGrid.TextMatrix(0, 3) = "Change"
    grdGrid.ColDataType(3) = flexDTBoolean
    grdGrid.Rows = 1
End Sub

Private Sub grdGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    lblInfo.Visible = True
    ActiveReadServer1 "Select Ave_Cost,Landed_Cost from Products where Product_code = '" & grdGrid.TextMatrix(NewRow, 0) & "'"
    If rs1.RecordCount > 0 Then
        lblInfo.Caption = " Average Cost: " & Format(rs1.Fields("Ave_Cost"), "0.00") & Chr$(13) & "  Landed Cost: " & Format(rs1.Fields("Landed_Cost"), "0.00") & " "
    End If
    rs1.Close
    Select Case NewCol
        Case 1
            grdGrid.Editable = flexEDKbdMouse
            grdGrid.ComboList = "..."
        Case 3
            grdGrid.Editable = flexEDKbdMouse
            grdGrid.ComboList = ""
        Case Else
            grdGrid.Editable = flexEDNone
    End Select
End Sub
Private Sub grdGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Load frmLabel
    frmLabel.Tag = "PriceChangeRow"
    DoEvents
    frmLabel.Show vbModal
End Sub
Private Sub grdGrid_EnterCell()
    If grdGrid.Col = 3 Then
        grdGrid.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub txtGross_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtMarkup.SetFocus
            KeyCode = 0
        Case 13, 40
            txtSellExcl.SetFocus
            KeyCode = 0
    End Select
End Sub
Private Sub txtMarkup_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtSellIncl.SetFocus
            KeyCode = 0
        Case 13, 40
            txtGross.SetFocus
            KeyCode = 0
    End Select
End Sub
Private Sub txtSellExcl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtGross.SetFocus
            KeyCode = 0
        Case 13, 40
            cmbTax.SetFocus
            KeyCode = 0
    End Select
End Sub

Private Sub txtSellExcl_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub cmbTax_Change()
    On Error Resume Next
    txtSellIncl.Tag = "1"
    If cmbTax.Text <> "" Then
        Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
    End If
    txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
    txtSellIncl.Tag = ""
    On Error GoTo 0
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    ActiveUpdateServer "Update Products set Selling_Price = " & txtSellIncl.Text & " where Product_Code = '" & txtProductCode.Text & "'"
    For i = 1 To grdGrid.Rows - 1
        If grdGrid.ValueMatrix(i, 3) = -1 Then
                ActiveUpdateServer "Update Products set Selling_Price = " & txtSellIncl.Text & " where Product_Code = '" & grdGrid.TextMatrix(i, 0) & "'"
        End If
    Next i
    Unload Me
End Sub
Private Sub txtMarkup_LostFocus()
    If txtMarkup.Text = "" Then txtMarkup.Text = "0"
End Sub
Private Sub txtSellExcl_LostFocus()
    If txtSellExcl.Text = "" Then txtSellExcl.Text = "0.00"
    txtSellExcl.Text = Format(txtSellExcl.Text, "0.00")
End Sub
Private Sub txtSellIncl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            cmbTax.SetFocus
            KeyCode = 0
        Case 13, 40
            cmdOk.SetFocus
            KeyCode = 0
    End Select
End Sub

Private Sub txtSellIncl_KeyPress(KeyAscii As Integer)
    If InStr(txtSellIncl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtSellIncl_LostFocus()
    If txtSellIncl.Text = "" Then txtSellIncl.Text = "0.00"
    txtSellIncl.Text = Format(txtSellIncl.Text, "0.00")
End Sub
Private Sub Form_Activate()
    Select Case Me.Tag
        Case "CostVar"
            cmbTax.Clear
            ActiveReadServer "Select * from Tax_Rates order by Tax_Type"
            While Not rs.EOF
                cmbTax.AddItem rs.Fields("Tax_type") & " - " & rs.Fields("Tax_Rate") & "% " & rs.Fields("Description")
                rs.MoveNext
            Wend
            rs.Close
            If cmbTax.ListCount > 1 Then
            cmbTax.Text = cmbTax.List(0)
            End If
            With frmReports
               ActiveReadServer2 "Select Product_Code,Landed_Cost,Ave_Cost,Sales_Tax,Tax_Type,Selling_Price,CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure " & _
               " ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) + Unit_of_Measure END as Description,(Select Description from Tax_Rates where Products.Tax_Type = Tax_Rates.Tax_Type) as Tax_Desc from Products where Product_Code ='" & .grdMain.TextMatrix(.grdMain.Row, 0) & "'"
               If rs2.RecordCount > 0 Then
                    txtProductCode.Text = rs2.Fields("Product_Code")
                    txtDescription.Text = rs2.Fields("Description")
                    lblInfo.Visible = False
                    lblAve.Visible = False
                    txtLandCost.Text = Format(rs2.Fields("Landed_Cost"), "0.00")
                    txtLandCost.ToolTipText = " Average Cost: " & Format(rs2.Fields("Ave_Cost"), "0.00") & " "
                    cmbTax.Text = rs2.Fields("Tax_Type") & " - " & rs2.Fields("Sales_Tax") & "% " & rs2.Fields("Tax_Desc")
                    txtSellIncl.Text = Format(rs2.Fields("Selling_Price"), "0.00")
               End If
               rs2.Close
            End With
        Case "ProdAN"
            cmbTax.Clear
            ActiveReadServer "Select * from Tax_Rates order by Tax_Type"
            While Not rs.EOF
                cmbTax.AddItem rs.Fields("Tax_type") & " - " & rs.Fields("Tax_Rate") & "% " & rs.Fields("Description")
                rs.MoveNext
            Wend
            rs.Close
            If cmbTax.ListCount > 1 Then
                cmbTax.Text = cmbTax.List(0)
            End If
            With frmReports
               ActiveReadServer2 "Select *,(Select Description from Tax_Rates where Products.Tax_Type = Tax_Rates.Tax_Type) as Tax_Desc from Products where Product_Code ='" & .grdMain.TextMatrix(.grdMain.Row, 0) & "'"
               If rs2.RecordCount > 0 Then
                    txtProductCode.Text = rs2.Fields("Product_Code")
                    txtDescription.Text = rs2.Fields("Description")
                    lblInfo.Visible = False
                    lblAve.Visible = False
                    txtLandCost.Text = Format(rs2.Fields("Ave_Cost"), "0.00")
                    txtLandCost.ToolTipText = " Average Cost: " & Format(rs2.Fields("Ave_Cost"), "0.00") & " "
                    cmbTax.Text = rs2.Fields("Tax_Type") & " - " & rs2.Fields("Sales_Tax") & "% " & rs2.Fields("Tax_Desc")
                    txtSellIncl.Text = Format(rs2.Fields("Selling_Price"), "0.00")
               End If
               rs2.Close
            End With
        Case "GRV"
            cmbTax.Clear
            ActiveReadServer "Select * from Tax_Rates order by Tax_Type"
            While Not rs.EOF
                cmbTax.AddItem rs.Fields("Tax_type") & " - " & rs.Fields("Tax_Rate") & "% " & rs.Fields("Description")
                rs.MoveNext
            Wend
            rs.Close
            If cmbTax.ListCount > 1 Then
                cmbTax.Text = cmbTax.List(0)
            End If
            With frmGRV
               ActiveReadServer2 "Select *,(Select Description from Tax_Rates where Products.Tax_Type = Tax_Rates.Tax_Type) as Tax_Desc from Products where Product_Code ='" & .grdGRV.TextMatrix(.grdGRV.Row, 0) & "'"
               If rs2.RecordCount > 0 Then
                    txtProductCode.Text = rs2.Fields("Product_Code")
                    txtDescription.Text = rs2.Fields("Description")
                    lblInfo.Caption = frmGRV.lblInfo.Tag
                    lblAve.Caption = frmGRV.lblAve.Tag
                    lblAve.ForeColor = frmGRV.lblAve.ForeColor
                    txtLandCost.Text = Format(.grdGRV.TextMatrix(.grdGRV.Row, 7), "0.00")
                    txtLandCost.ToolTipText = " Average Cost: " & Format(rs2.Fields("Ave_Cost"), "0.00") & " "
                    cmbTax.Text = rs2.Fields("Tax_Type") & " - " & rs2.Fields("Sales_Tax") & "% " & rs2.Fields("Tax_Desc")
                    txtSellIncl.Text = Format(rs2.Fields("Selling_Price"), "0.00")
               End If
               rs2.Close
            End With
    End Select
    txtSellIncl.SelStart = Len(txtSellIncl.Text)
End Sub
Private Sub txtGross_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtGross_Change()
    On Error Resume Next
    If txtGross.Tag <> "" Then Exit Sub
    txtSellExcl.Tag = "1"
    txtSellExcl.Text = Format(txtLandCost.Text / ((100 - Val(txtGross.Text)) / 100), "0.00")
    txtSellExcl.Tag = ""
    If Val(txtLandCost.Text) <> 0 Then
        txtMarkup.Tag = "1"
        txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
        txtMarkup.Tag = ""
    End If
    txtSellIncl.Tag = "1"
    Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
    txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
    txtSellIncl.Tag = ""
    On Error GoTo 0
End Sub
Private Sub txtGross_LostFocus()
    If txtGross.Text = "" Then txtGross.Text = "0"
End Sub
Private Sub txtMarkup_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtMarkup_Change()
    On Error Resume Next
    If txtMarkup.Tag <> "" Then Exit Sub
    If Val(txtLandCost.Text) <> 0 Then
        txtSellExcl.Tag = "1"
        txtSellExcl.Text = Format(Val(txtLandCost.Text) * ((100 + Val(txtMarkup.Text)) / 100), "0.00")
        txtSellExcl.Tag = ""
    End If
    If Val(txtSellExcl) <> 0 Then
        txtGross.Tag = "1"
        txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl) * 100), 3)
        txtGross.Tag = ""
    Else
        txtGross.Tag = "1"
        txtGross.Text = "0"
        txtGross.Tag = ""
    End If
    txtSellIncl.Tag = "1"
    Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
    If Val(txtLandCost.Text) <> 0 Then
        txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
    End If
    txtSellIncl.Tag = ""
    On Error GoTo 0
End Sub
Private Sub txtSellExcl_Change()
    On Error Resume Next
    If txtSellExcl.Tag <> "" Then Exit Sub
    If Val(txtLandCost.Text) <> 0 Then
        txtMarkup.Tag = "1"
        txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost) * 100), 3)
        txtMarkup.Tag = ""
    End If
    If Val(txtSellExcl) <> 0 Then
        txtGross.Tag = "1"
        txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl) * 100), 3)
        txtGross.Tag = ""
    Else
        txtGross.Tag = "1"
        txtGross.Text = "0"
        txtGross.Tag = ""
    End If
    txtSellIncl.Tag = "1"
    Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
    If txtSellExcl.Text <> "N/A" Then
        txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
    End If
    txtSellIncl.Tag = ""
    On Error GoTo 0
End Sub
Private Sub txtSellIncl_Change()
    On Error Resume Next
    If txtSellIncl.Tag <> "" Then Exit Sub
    Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
    If txtSellIncl.Text <> "N/A" Then
        txtSellExcl.Tag = "1"
        txtSellExcl.Text = Format(txtSellIncl.Text / ((100 + Tax) / 100), "0.00")
        txtSellExcl.Tag = ""
    End If
    txtSellExcl.Tag = ""
    If Val(txtLandCost.Text) <> 0 Then
        txtMarkup.Tag = "1"
        txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
        txtMarkup.Tag = ""
    End If
    If Val(txtSellExcl) <> 0 Then
        txtGross.Tag = "1"
        txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
        txtGross.Tag = ""
    Else
        txtGross.Tag = "1"
        txtGross.Text = "0"
        txtGross.Tag = ""
    End If
    On Error GoTo 0
End Sub
