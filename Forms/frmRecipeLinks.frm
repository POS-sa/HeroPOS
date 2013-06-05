VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRecipeLinks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Recipe Links..."
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUnit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   11040
      TabIndex        =   8
      Top             =   240
      Width           =   2445
   End
   Begin VB.TextBox txtLanded 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6150
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   3315
   End
   Begin VB.TextBox txtProdCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   3315
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   7815
   End
   Begin VSFlex8Ctl.VSFlexGrid grdRecipe 
      Height          =   6000
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   13575
      _cx             =   23945
      _cy             =   10583
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
      BackColorSel    =   16703960
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRecipeLinks.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   5
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
      WordWrap        =   -1  'True
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
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   345
      Left            =   12450
      TabIndex        =   0
      Top             =   7080
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Close"
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
   Begin btButtonEx.ButtonEx cmdUpdate 
      Height          =   345
      Left            =   60
      TabIndex        =   12
      Top             =   7080
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   609
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Update Changes"
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
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unit of Measure: "
      Height          =   195
      Left            =   9510
      TabIndex        =   11
      Top             =   630
      Width           =   1335
   End
   Begin MSForms.ComboBox cmbUnit 
      Height          =   315
      Left            =   10860
      TabIndex        =   10
      Tag             =   "Up"
      Top             =   570
      Width           =   2655
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4683;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Size: "
      Height          =   195
      Left            =   9510
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin MSForms.Image Image2 
      Height          =   345
      Left            =   10860
      Top             =   150
      Width           =   2655
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "4683;609"
   End
   Begin MSForms.Image Image1 
      Height          =   345
      Left            =   6060
      Top             =   150
      Width           =   3435
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "6059;609"
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Landed Cost: "
      Height          =   195
      Left            =   4740
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code: "
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description: "
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   570
      Width           =   1335
   End
   Begin MSForms.Image Image11 
      Height          =   345
      Left            =   1530
      Top             =   150
      Width           =   3435
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "6059;609"
   End
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   1530
      Top             =   540
      Width           =   7965
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "14049;609"
   End
   Begin MSForms.Image Image4 
      Height          =   915
      Left            =   60
      Top             =   60
      Width           =   13575
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "23945;1614"
   End
End
Attribute VB_Name = "frmRecipeLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbUnit_Change()
    Cmdupdate.Enabled = True
End Sub
Private Sub cmbUnit_LostFocus()
    Recalc
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdUpdate_Click()
    ActiveUpdateServer "Update Products set Unit_Size = '" & txtUnit.Text & "',Unit_of_Measure = '" & cmbUnit.Text & "' where Product_Code ='" & txtProdCode.Text & "'"
    DoEvents
    For i = 1 To grdRecipe.Rows - 1
        ActiveUpdateServer "Update Recipes set Description = '" & grdRecipe.TextMatrix(i, 1) & "',Unit_of_Measure='" & grdRecipe.TextMatrix(i, 2) & "',Qty_Used='" & grdRecipe.TextMatrix(i, 3) & "' where Line_No = " & grdRecipe.TextMatrix(i, 6)
    Next i
    MsgBox "Recipe Updated Successfully.", vbInformation, "HeroPOS"
    Unload Me
End Sub

Private Sub Form_Load()
    cmbUnit.Clear
    cmbUnit.AddItem "ml"
    cmbUnit.AddItem "lt"
    cmbUnit.AddItem "g"
    cmbUnit.AddItem "kg"
    cmbUnit.AddItem "ton"
    cmbUnit.AddItem "each"
    cmbUnit.AddItem "box"
    cmbUnit.AddItem "Preparation Recipe"
    cmbUnit.Text = "each"
    grdRecipe.Rows = 1
    grdRecipe.TextMatrix(0, 0) = "Recipe Line"
    grdRecipe.TextMatrix(0, 1) = "Stock Item Used"
    grdRecipe.TextMatrix(0, 2) = "Unit of Measure"
    grdRecipe.TextMatrix(0, 3) = "Qty Used"
    grdRecipe.TextMatrix(0, 4) = "Landed Cost"
    grdRecipe.TextMatrix(0, 5) = "Used In"
    grdRecipe.ColWidth(0) = grdRecipe.Width * 0.15
    grdRecipe.ColWidth(1) = grdRecipe.Width * 0.25
    grdRecipe.ColWidth(2) = grdRecipe.Width * 0.1
    grdRecipe.ColWidth(3) = grdRecipe.Width * 0.1
    grdRecipe.ColWidth(4) = grdRecipe.Width * 0.1
    grdRecipe.ColWidth(5) = grdRecipe.Width * 0.2
    grdRecipe.ColAlignment(0) = flexAlignLeftCenter
    grdRecipe.ColAlignment(1) = flexAlignLeftCenter
    grdRecipe.ColAlignment(3) = flexAlignRightCenter
    grdRecipe.ColAlignment(4) = flexAlignRightCenter
    grdRecipe.ColHidden(6) = True
    txtProdCode.Text = frmProducts.txtProductCode.Text
    txtDescription.Text = frmProducts.grdProd.TextMatrix(frmProducts.grdProd.Row, 1)
    txtLanded.Text = Format(frmProducts.grdProd.TextMatrix(frmProducts.grdProd.Row, 4), "0.00")
    txtUnit = frmProducts.txtUnitSize
    cmbUnit = frmProducts.cmbUnit.Text
    grdRecipe.Rows = 1
    ActiveReadServer "Select * from Recipes where Line_Code= '" & txtProdCode.Text & "' order by Line_No"
    While Not rs.EOF
        grdRecipe.Rows = grdRecipe.Rows + 1
        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1) = rs.Fields("Description")
        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs.Fields("Product_Code")
        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs.Fields("Line_No")
        ActiveReadServer1 "Select Description from Products where Product_Code ='" & rs.Fields("Product_Code") & "'"
        If rs1.RecordCount > 0 Then
            grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs.Fields("Product_Code") & " - " & rs1.Fields("Description")
        End If
        rs1.Close
        If rs.Fields("Unit_of_Measure") <> cmbUnit.Text Then
            Select Case UCase(cmbUnit.Text & " to " & rs.Fields("Unit_of_Measure"))
                Case "ML TO SINGLE TOT", "ML TO DOUBLE TOT", "ML TO LT", "LT TO ML", "KG TO G", "G TO KG", "ML TO ML", "G TO G", "LT TO LT", "KG TO KG"
                    If txtUnit.Text = "" Then
                        txtUnit.BackColor = vbRed
                        Image2.BackColor = vbRed
                        txtUnit.ForeColor = vbWhite
                    End If
                Case "EACH TO EACH"
                Case Else
                    If txtUnit.Text = "" Then
                        txtUnit.BackColor = vbRed
                        Image2.BackColor = vbRed
                        txtUnit.ForeColor = vbWhite
                    End If
                    grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 2) = vbRed
                    grdRecipe.Cell(flexcpForeColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 2) = vbWhite
            End Select
        Else
            If txtUnit.Text = "" Then
                If cmbUnit.Text <> "each" Then
                    txtUnit.BackColor = vbRed
                    Image2.BackColor = vbRed
                    txtUnit.ForeColor = vbWhite
                End If
            End If
        End If
        Select Case rs.Fields("Line_Type")
            Case 0
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Message"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
            Case 1
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Preparation Recipe"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
            Case 2
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Sales Item"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
            Case 3
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
            Case 4
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item (Hidden)"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
            Case 5
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Price/Size Change"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = "0.00"
                grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
            Case 6
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Sales Item (Choice)"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
            Case 7
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item (Choice)"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
            Case 8
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Exit"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = " "
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = " "
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = " "
                grdRecipe.MergeRow(grdRecipe.Rows - 1) = True
                grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 2, grdRecipe.Rows - 1, 4) = &HE0E0E0
        End Select
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub grdRecipe_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Recalc
End Sub
Private Sub grdRecipe_EnterCell()
    grdRecipe.Editable = flexEDNone
    If grdRecipe.Col = 2 Then
        grdRecipe.Editable = flexEDKbdMouse
    End If
    Select Case cmbUnit.Text
        Case "ml"
            grdRecipe.ColComboList(2) = ""
            grdRecipe.ColComboList(2) = "ml|Single Tot|Double Tot"
            grdRecipe.TextMatrix(grdRecipe.Row, 2) = "ml"
        Case "lt"
            grdRecipe.ColComboList(2) = ""
            grdRecipe.ColComboList(2) = "ml|lt|Single Tot|Double Tot"
            grdRecipe.TextMatrix(grdRecipe.Row, 2) = "ml"
        Case "g"
            grdRecipe.ColComboList(2) = ""
            grdRecipe.TextMatrix(grdRecipe.Row, 2) = "g"
            grdRecipe.Editable = flexEDNone
        Case "kg"
            grdRecipe.ColComboList(2) = ""
            grdRecipe.ColComboList(2) = "g|kg"
            grdRecipe.TextMatrix(grdRecipe.Row, 2) = "g"
        Case Else
            grdRecipe.ColComboList(2) = ""
            grdRecipe.TextMatrix(grdRecipe.Row, 2) = "each"
    End Select
    If grdRecipe.Col = 3 Then
        grdRecipe.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub grdRecipe_KeyDown(KeyCode As Integer, Shift As Integer)
    If grdRecipe.Col = 3 Then
        Select Case KeyCode
            Case 8, 46, 48 To 57
                grdRecipe.EditCell
            Case 13, 37, 38, 39, 40
            Case Else: KeyCode = 0
        End Select
    End If
End Sub

Private Sub grdRecipe_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If grdRecipe.Col = 3 Then
        If InStr(grdRecipe.EditText, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
        Select Case KeyAscii
            Case 8, 46, 48 To 57
            Case Else: KeyAscii = 0
        End Select
    End If
End Sub
Private Sub txtUnit_Change()
    Cmdupdate.Enabled = True
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtUnit_LostFocus()
    Recalc
End Sub
Private Sub Recalc()
    If cmbUnit.Text = "each" Then txtUnit.Text = ""
    For i = 1 To grdRecipe.Rows - 1
        ActiveReadServer "Select Description from Products where Product_Code = '" & txtProdCode.Text & "'"
        If rs.RecordCount > 0 Then
            grdRecipe.TextMatrix(i, 1) = rs.Fields("Description") & " " & txtUnit.Text & cmbUnit.Text & " ," & txtProdCode.Text
        End If
        rs.Close
    Next i
End Sub
