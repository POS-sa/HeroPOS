VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmProduction 
   Caption         =   " Production..."
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin btButtonEx.ButtonEx cmdRefresh 
      Height          =   325
      Left            =   5310
      TabIndex        =   9
      Top             =   750
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Refresh List..."
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
   Begin VB.TextBox txtLandCost 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   11835
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   330
      Width           =   1725
   End
   Begin VSFlex8Ctl.VSFlexGrid grdRecipe 
      Height          =   7680
      Left            =   6960
      TabIndex        =   0
      Top             =   1110
      Width           =   6675
      _cx             =   11774
      _cy             =   13547
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   15772582
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16381166
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   7680
      Left            =   60
      TabIndex        =   1
      Top             =   1110
      Width           =   6825
      _cx             =   12039
      _cy             =   13547
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16306895
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16381166
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label3 
      Caption         =   "Stock Items to Use."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7110
      TabIndex        =   8
      Top             =   780
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Products to be Produced."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Top             =   780
      Width           =   3135
   End
   Begin MSForms.Label lblCost 
      Height          =   225
      Left            =   10290
      TabIndex        =   6
      Top             =   330
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Landed Cost:"
      Size            =   "2461;397"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   3
      Left            =   11730
      Top             =   270
      Width           =   1845
      BackColor       =   16777215
      Size            =   "3254;529"
   End
   Begin MSForms.Image Image2 
      Height          =   8145
      Left            =   60
      Top             =   690
      Width           =   6855
      BackColor       =   14737632
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12091;14367"
   End
   Begin MSForms.Image Image4 
      Height          =   8145
      Index           =   0
      Left            =   6960
      Top             =   690
      Width           =   6705
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "11827;14367"
   End
   Begin VB.Label lblUsers 
      Caption         =   "Production Location."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   210
      TabIndex        =   4
      Top             =   150
      Width           =   3135
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   2
      Left            =   2940
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   1
      Left            =   2610
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   3
      Left            =   2280
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image5 
      Height          =   90
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   2115
      BackColor       =   16761024
      Size            =   "3731;159"
   End
   Begin VB.Label Label1 
      Caption         =   "Production Recipe."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7110
      TabIndex        =   3
      Top             =   150
      Width           =   3135
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   4
      Left            =   9840
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   5
      Left            =   9510
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   6
      Left            =   9180
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image6 
      Height          =   90
      Left            =   7020
      Top             =   480
      Width           =   2115
      BackColor       =   16761024
      Size            =   "3731;159"
   End
   Begin MSForms.ComboBox cmbTransLoc 
      Height          =   375
      Left            =   3450
      TabIndex        =   2
      Top             =   180
      Width           =   3405
      DisplayStyle    =   7
      Size            =   "6006;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image1 
      Height          =   555
      Left            =   60
      Top             =   90
      Width           =   6855
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12091;979"
   End
   Begin MSForms.Image Image3 
      Height          =   555
      Left            =   6960
      Top             =   90
      Width           =   6705
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "11827;979"
   End
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTransLoc_Change()
    On Error Resume Next
    grdMain.Rows = 1
    If cmbTransLoc.Text <> "<Select a Location>" Then
        If Trim(cmbTransLoc.Text) <> "" Then
            ActiveReadServer "Select * From Production_View where Location_No = " & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & " order by Description"
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_code")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description")
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = Round(Val(rs.Fields("Stock_on_Hand") & ""), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = "0"
                grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Department_No")
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = rs.Fields("Unit_of_Measure")
                rs.MoveNext
            Wend
            rs.Close
            If grdMain.Rows > 1 Then
                grdMain.Row = 1
            End If
        End If
        grdMain.Enabled = True
        grdRecipe.Enabled = True
    Else
        grdMain.Enabled = False
        grdRecipe.Enabled = False
    End If
    On Error GoTo 0
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    grdMain.SetFocus
    grdMain.Rows = 1
    DoEvents
    ActiveReadServer "Select * From Production_View where Location_No = " & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & " order by Description"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_code")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Round(Val(rs.Fields("Stock_on_Hand") & ""), 3)
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = "0"
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Department_No")
        rs.MoveNext
    Wend
    rs.Close
    If grdMain.Rows > 1 Then
        grdMain.Row = 1
    End If
    grdMain.Enabled = True
    grdRecipe.Enabled = True
    cmdRefresh.Enabled = False
    On Error GoTo 0
End Sub
Private Sub Form_Load()
    frmMain.Toolbar1.Buttons(2).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    cmbTransLoc.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    cmbTransLoc.AddItem "<Select a Location>"
    While Not rs.EOF
        cmbTransLoc.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmbTransLoc.Text = "<Select a Location>"
    grdMain.Rows = 1
    grdMain.Cols = 6
    grdMain.TextMatrix(0, 0) = "Product Code"
    grdMain.TextMatrix(0, 1) = "Description"
    grdMain.TextMatrix(0, 2) = "Stock on Hand"
    grdMain.TextMatrix(0, 3) = " Production Qty"
    grdMain.RowHeight(0) = 550
    grdRecipe.RowHeight(0) = 550
    grdMain.ColWidth(0) = grdMain.Width * 0.2
    grdMain.ColWidth(1) = grdMain.Width * 0.45
    grdMain.ColWidth(2) = grdMain.Width * 0.15
    grdMain.ColWidth(3) = grdMain.Width * 0.2
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignRightCenter
    grdMain.ColAlignment(3) = flexAlignRightCenter
    grdMain.ColHidden(4) = True
    grdMain.ColHidden(5) = True
    grdRecipe.Rows = 1
    grdRecipe.TextMatrix(0, 0) = "Recipe Line"
    grdRecipe.TextMatrix(0, 1) = "Description or Product"
    grdRecipe.TextMatrix(0, 2) = "Unit of Measure"
    grdRecipe.TextMatrix(0, 3) = "Qty Used"
    grdRecipe.TextMatrix(0, 4) = "Landed Cost"
    grdRecipe.ColWidth(0) = grdRecipe.Width * 0.2
    grdRecipe.ColWidth(1) = grdRecipe.Width * 0.35
    grdRecipe.ColWidth(2) = grdRecipe.Width * 0.15
    grdRecipe.ColAlignment(0) = flexAlignLeftCenter
    grdRecipe.ColAlignment(1) = flexAlignLeftCenter
    grdRecipe.ColAlignment(3) = flexAlignRightCenter
    grdRecipe.ColWidth(3) = grdRecipe.Width * 0.15
    grdRecipe.ColWidth(4) = grdRecipe.Width * 0.12
    grdRecipe.ColAlignment(4) = flexAlignRightCenter
    grdRecipe.ColHidden(5) = True
    grdRecipe.ColHidden(6) = True
    grdRecipe.ColHidden(7) = True
    grdRecipe.Rows = 2
    frmMain.Toolbar1.Buttons(2).Caption = "Produce"
End Sub

Public Sub Produce()

        ActiveReadServer "Select (Select isnull(Max(convert(int,Production_No)),0)+1 from Production_Journal) as Production_No"
        If rs.RecordCount > 0 Then
            ProductionNo = rs.Fields("Production_No")
        End If
        rs.Close
        For i = 1 To grdMain.Rows - 1
            grdMain.Row = i
                 If grdMain.ValueMatrix(i, 3) <> 0 Then
                    ActiveReadServer "Select Stock_on_Hand from Quantities where Product_Code = '" & grdMain.TextMatrix(i, 0) & "' and Location_No = " & Trim(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1))
                    If rs.RecordCount > 0 Then
                        ActiveUpdateServer "Update Quantities Set Stock_on_Hand = Stock_on_Hand + " & Val(grdMain.ValueMatrix(i, 3)) & " where Product_Code = '" & grdMain.ValueMatrix(i, 0) & "' and Location_No = " & Trim(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1))
                    Else
                        ActiveUpdateServer "INSERT INTO Quantities (Product_Code,Location_No,Stock_on_Hand) values ('" & grdMain.TextMatrix(i, 0) & "'," & Trim(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & "," & Val(grdMain.ValueMatrix(i, 3)) & ")"
                    End If
                    rs.Close
                    For b = 1 To grdRecipe.Rows - 1
                        If grdMain.ValueMatrix(i, 0) = grdRecipe.ValueMatrix(b, 8) Then
                            ActiveReadServer1 "Select Unit_Size,Unit_of_Measure from Products where Product_Code = '" & Mid(grdRecipe.TextMatrix(b, 1), InStrRev(grdRecipe.TextMatrix(b, 1), ",") + 1) & "'"
                            UnitSize = rs1.Fields("Unit_Size")
                            If rs1.Fields("Unit_of_Measure") <> grdRecipe.TextMatrix(b, 2) Then
                                Select Case UCase(rs1.Fields("Unit_of_Measure") & " to " & grdRecipe.TextMatrix(b, 2))
                                    Case "ML TO LT"
                                        UnitSize = rs1.Fields("Unit_Size") / 1000
                                    Case "LT TO ML"
                                        If rs1.Fields("Unit_Size") = 0 Then
                                            UnitSize = 1000
                                        Else
                                            UnitSize = rs1.Fields("Unit_Size") * 1000
                                        End If
                                    Case "G TO KG"
                                        UnitSize = rs1.Fields("Unit_Size") / 1000
                                    Case "KG TO G"
                                        If rs1.Fields("Unit_Size") = 0 Then
                                            UnitSize = 1000
                                        Else
                                            UnitSize = rs1.Fields("Unit_Size") * 1000
                                        End If
                                    Case Else
                                        UnitSize = rs1.Fields("Unit_Size")
                                End Select
                            End If
                            If Val(UnitSize & "") <> 0 Then
                                Select Case grdRecipe.TextMatrix(b, 3)
                                    Case Is = "1 x 25ml"
                                        Qty = Round(25 / UnitSize, 4)
                                    Case Is = "2 x 25ml"
                                        Qty = Round(50 / UnitSize, 4)
                                    Case Else
                                        Qty = Round(grdRecipe.TextMatrix(b, 3) / UnitSize, 4)
                                 End Select
                                
                            Else
                                Qty = grdRecipe.ValueMatrix(b, 3)
                            End If
                            rs1.Close
                            
                            ActiveUpdateServer "INSERT INTO [Production_Journal](" & _
                                    "[Product_Code], " & _
                                    "[Location_No],  " & _
                                    "[Ave_Cost],  " & _
                                    "[Qty_Produced],  " & _
                                    "[UOM_Produced], " & _
                                    "[Date_Time], " & _
                                    "[Product_Code_Consumed]," & _
                                    "[Qty_Consumed]," & _
                                    "[UOM_Consumed], " & _
                                    "[User_no])" & _
                                " VALUES('" & _
                                    grdMain.TextMatrix(i, 0) & "'," & _
                                    Trim(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & "," & _
                                    txtLandCost.Text / grdMain.ValueMatrix(i, 3) & "," & _
                                    Val(grdMain.ValueMatrix(i, 3)) & ",'" & _
                                    grdMain.TextMatrix(i, 5) & "', " & _
                                    "getdate() ,'" & _
                                    Mid(grdRecipe.TextMatrix(b, 1), InStrRev(grdRecipe.TextMatrix(b, 1), ",") + 1) & "','" & _
                                    Qty & "','" & _
                                    grdRecipe.TextMatrix(b, 2) & "', " & _
                                    UserRecord.User_Number & _
                                ")"
                            
                            ActiveReadServer "Select Stock_on_Hand from Quantities where Product_Code = '" & Mid(grdRecipe.TextMatrix(b, 1), InStrRev(grdRecipe.TextMatrix(b, 1), ",") + 1) & "' and Location_No = " & Trim(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1))
                            If rs.RecordCount > 0 Then
                                ActiveUpdateServer "Update Quantities Set Stock_on_Hand = Stock_on_Hand - " & Qty & " where Product_Code = '" & Mid(grdRecipe.TextMatrix(b, 1), InStrRev(grdRecipe.TextMatrix(b, 1), ",") + 1) & "' and Location_No = " & Trim(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1))
                            Else
                                ActiveUpdateServer "INSERT INTO Quantities (Product_Code,Location_No,Stock_on_Hand) values ('" & grdRecipe.TextMatrix(b, 0) & "'," & Trim(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & "," & Qty * -1 & ")"
                            End If
                            rs.Close
                        End If
                    Next b
                End If
        Next i
        
        Exit Sub
        On Error GoTo trap
        Screen.MousePointer = 11
        PrintErr = 0
        Slip_Port = ""
        filenum = FreeFile
        Close #filenum
        If Slip_Printer = "<None>" Then
            GoTo far
        End If
        If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
            If InStr(Trim(Slip_Printer), "\\") = 0 Then
                If Slip_Port = "" Then
                    Open "\\" & Comp_Name & "\" & Slip_Printer For Output As filenum
                Else
                    Open Slip_Port For Output As filenum
                End If
            Else
                If Slip_Port = "" Then
                    Open Slip_Printer For Output As filenum
                Else
                    Open Slip_Port For Output As filenum
                End If
            End If
            If Slip_Port <> "" Then
                If UCase(Left(Slip_Port, 2)) = "NE" Then
                    Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
                Else
                    Open Slip_Port For Output As filenum
                End If
            End If
        Else
            Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
        End If
        Print #filenum, Chr(27) & Chr(64);
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, Chr(27) & Chr(97) & Chr(49);
        Print #filenum, Chr(27) & Chr(33) & Chr(16);
        Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, UCase(Branch_Name)
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        Print #filenum, Chr(27) & Chr(69) & Chr(48);
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, Chr(27) & Chr(97) & Chr(49);
        Print #filenum, Chr(27) & Chr(33) & Chr(16);
        Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, "STOCK PRODUCTION"
        Print #filenum, Format(Date, "YYYY-MM-DD") & " " & Time
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        Print #filenum, Chr(27) & Chr(69) & Chr(48);
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
        Print #filenum, Chr(27) & Chr(33) & Chr(16);
        Print #filenum, Chr(27) & Chr(97) & Chr(49);
        Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, UCase(cmdUser.Caption)
        Print #filenum, Chr(27) & Chr(69) & Chr(48);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        Print #filenum, String(33, "=")
        Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, Chr(27) & Chr(97) & Chr(50);
        Print #filenum, Chr(27) & Chr(69) & Chr(48);
        Print #filenum, Chr(27) & Chr(77) & Chr(48);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        Print #filenum, Chr(27) & Chr(97) & Chr(48);
        Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        '
        For i = 1 To grdMain.Rows - 1
            If grdMain.ValueMatrix(i, 3) <> 0 Then
            End If
        Next i
        Print #filenum, String(40, "-")
        Print #filenum, Chr(27) & Chr(97) & Chr(48);
        Print #filenum, "PRODUCED BY:"
        Print #filenum, ""
        Print #filenum, ""
        Print #filenum, ""
        Print #filenum, "ACCEPTED BY:"
        Print #filenum, ""
        Print #filenum, ""
        Print #filenum, ""
        Print #filenum, Chr(27) & Chr(50);
        Print #filenum, String(40, "=")
        Print #filenum, Chr(27) & Chr(100) & Chr(7);
        Print #filenum, Chr(29) & "V" & Chr(49);
        Print #filenum, Chr(27) & Chr(64);
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
        Close #filenum
        Screen.MousePointer = 0
far:
        On Error GoTo 0
        
        Exit Sub
trap:
        If PrintErr = 0 Then
            PrintErr = 1
            Dim x As Printer
            For Each x In Printers
                If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                    Slip_Port = x.Port
                    Exit For
                End If
            Next
            Resume Next
        End If
        Load frmError
        frmStockTake.Tag = "Not Now"
        frmError.Caption = " Printer Error - " & Slip_Printer
        frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
        frmError.lblError.Caption = err.Description
        DoEvents
        frmError.Show vbModal
        Screen.MousePointer = 0
        Unload Me
    
    On Error GoTo 0

End Sub
Public Sub SaveProduction()
    ActiveUpdateServer "Delete from Production where Product_Code='" & grdMain.TextMatrix(grdMain.Row, 0) & "'"
    DoEvents
    For i = 1 To grdRecipe.Rows - 1
        If grdRecipe.TextMatrix(i, 1) <> "" Then
            Select Case grdRecipe.TextMatrix(i, 0)
                Case "Message"
                    LineType = 0
                Case "Stock Item"
                    LineType = 3
                Case "Stock Item (Hidden)"
                    LineType = 4
            End Select
            ActiveUpdateServer "INSERT INTO Production ([Line_Type], [Product_Code], [Line_Code], [Description], [Unit_of_Measure], [Qty_Used], [Cost])" & _
            "VALUES (" & LineType & ",'" & grdMain.TextMatrix(grdMain.Row, 0) & "'," & Trim(Val(Trim(Mid(grdRecipe.TextMatrix(i, 1), InStrRev(grdRecipe.TextMatrix(i, 1), ",") + 1)))) & ",'" & grdRecipe.TextMatrix(i, 1) & "','" & Trim(grdRecipe.TextMatrix(i, 2)) & "','" & Trim(grdRecipe.TextMatrix(i, 3)) & "'," & Val(grdRecipe.TextMatrix(i, 4)) & ")"
        End If
    Next i
    MsgBox "Production recipe update successful", vbInformation, "HeroPOS"
    TotalProd = 0
    For i = 1 To grdMain.Rows - 1
        TotalProd = TotalProd + grdMain.ValueMatrix(i, 3)
    Next i
    If TotalProd = 0 Then
        frmMain.Toolbar1.Buttons(2).Enabled = False
    Else
        frmMain.Toolbar1.Buttons(2).Enabled = True
    End If
    frmMain.Toolbar1.Buttons(4).Enabled = False
    grdMain.HighLight = flexHighlightWithFocus
    grdMain.SelectionMode = flexSelectionFree
    grdMain.Enabled = True
    grdMain.SetFocus
End Sub

Private Sub grdMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    TotalProd = 0
    For i = 1 To grdMain.Rows - 1
        TotalProd = TotalProd + grdMain.ValueMatrix(i, 3)
    Next i
    If TotalProd = 0 Then
        frmMain.Toolbar1.Buttons(2).Enabled = False
    Else
        frmMain.Toolbar1.Buttons(2).Enabled = True
    End If
    TotalProd = 0
    For i = 1 To grdRecipe.Rows - 1
        TotalProd = TotalProd + grdRecipe.ValueMatrix(i, 3)
    Next i
    If grdMain.ValueMatrix(grdMain.Row, 3) <> 0 Then
        If TotalProd <> 0 Then
            frmMain.Toolbar1.Buttons(2).Enabled = True
        Else
            frmMain.Toolbar1.Buttons(2).Enabled = False
        End If
    End If
End Sub

Private Sub grdMain_EnterCell()
    visible_rows = 0
    need_line = False
    For i = 1 To grdRecipe.Rows - 1
        
        If grdRecipe.TextMatrix(i, 8) <> "" Then
            If grdRecipe.TextMatrix(i, 8) = grdMain.TextMatrix(grdMain.Row, 0) Then
                grdRecipe.RowHidden(i) = False
                visible_rows = visible_rows + 1
            Else
                grdRecipe.RowHidden(i) = True
            End If
        Else
            visible_rows = visible_rows + 1
        End If
    Next i
    If visible_rows = 0 Then grdRecipe.Rows = grdRecipe.Rows + 1
    grdMain.Col = 3
End Sub

Private Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 45, 48 To 57, 96 To 105, 109, 110, 189
            Select Case grdMain.Col
                Case 3
                    grdMain.EditCell
            End Select
        Case Else
    End Select
End Sub
Private Sub grdMain_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(grdMain.EditText, ".") <> 0 And KeyAscii = 46 Then
        KeyAscii = 0
    End If
    Select Case KeyAscii
        Case 8, 13, 27, 45, 46, 48 To 57
        Case Else
            If Col = 3 Then KeyAscii = 0
    End Select
End Sub

Private Sub grdRecipe_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If grdRecipe.Tag = grdRecipe.Text Then Exit Sub
    grdMain.HighLight = flexHighlightAlways
    grdMain.SelectionMode = flexSelectionByRow
    frmMain.Toolbar1.Buttons(2).Enabled = False
    Select Case Col
        Case 0
            Select Case grdRecipe.TextMatrix(Row, 0)
                Case "Message", "Preparation Recipe", "Price/Size Change", "Exit"
                    grdRecipe.TextMatrix(Row, 1) = ""
                    grdRecipe.TextMatrix(Row, 2) = " "
                    grdRecipe.TextMatrix(Row, 3) = " "
                    grdRecipe.TextMatrix(Row, 4) = " "
                    grdRecipe.MergeRow(Row) = True
                    grdRecipe.Cell(flexcpBackColor, grdRecipe.Rows - 1, 0, grdRecipe.Rows - 1, 4) = &HC0FFFF
                    grdRecipe.Cell(flexcpFontBold, grdRecipe.Rows - 1, 1, grdRecipe.Rows - 1, 1) = True
                Case Else
                    grdRecipe.TextMatrix(Row, 1) = ""
                    grdRecipe.TextMatrix(Row, 2) = ""
                    grdRecipe.TextMatrix(Row, 3) = ""
                    grdRecipe.TextMatrix(Row, 4) = ""
                    grdRecipe.MergeRow(Row) = False
                    grdRecipe.Cell(flexcpBackColor, Row, 2, Row, 4) = grdRecipe.BackColor
            End Select
        Case 1
            Select Case grdRecipe.TextMatrix(Row, 0)
                Case "Message", "Price/Size Change", "Exit"
                    grdRecipe.TextMatrix(Row, 1) = UCase(Left(grdRecipe.TextMatrix(Row, 1), 1)) & Mid(grdRecipe.TextMatrix(Row, 1), 2)
                    grdRecipe.Cell(flexcpBackColor, Row, 0, Row, 4) = &HC0FFFF
                    grdRecipe.Cell(flexcpFontBold, Row, 1, Row, 1) = True
                Case "Preparation Recipe"
                Case "Sales Item", "Stock Item", "Stock Item (Hidden)", "Stock Item (Choice)", "Sales Item (Choice)"
                    If InStrRev(grdRecipe.TextMatrix(Row, 1), ",") = 0 Then
                        grdRecipe.TextMatrix(Row, 1) = ""
                    Else
                        grdRecipe.TextMatrix(Row, 8) = grdMain.TextMatrix(grdMain.Row, 0)
                        ActiveReadServer "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(Row, 1), InStrRev(grdRecipe.TextMatrix(Row, 1), ",") + 1) & "'"
                        If rs.RecordCount > 0 Then
                            Select Case rs.Fields("Unit_of_Measure")
                                Case "ml"
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.ColComboList(2) = "ml|Single Tot|Double Tot"
                                    grdRecipe.TextMatrix(Row, 2) = "ml"
                                    grdRecipe.TextMatrix(Row, 3) = rs.Fields("Unit_Size")
                                Case "lt"
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.ColComboList(2) = "ml|lt|Single Tot|Double Tot"
                                    grdRecipe.TextMatrix(Row, 2) = "ml"
                                    grdRecipe.TextMatrix(Row, 3) = rs.Fields("Unit_Size") * 1000
                                Case "g"
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.TextMatrix(Row, 2) = "g"
                                    grdRecipe.TextMatrix(Row, 3) = rs.Fields("Unit_Size")
                                Case "kg"
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.ColComboList(2) = "g|kg"
                                    grdRecipe.TextMatrix(Row, 2) = "g"
                                    grdRecipe.TextMatrix(Row, 3) = rs.Fields("Unit_Size") * 1000
                                Case Else
                                    grdRecipe.ColComboList(2) = ""
                                    grdRecipe.TextMatrix(Row, 2) = "each"
                                    grdRecipe.TextMatrix(Row, 3) = "1"
                            End Select
                        End If
                        grdRecipe.TextMatrix(Row, 5) = rs.Fields("Unit_Size") & ""
                        If grdRecipe.TextMatrix(Row, 5) = 0 Then grdRecipe.TextMatrix(Row, 5) = "1"
                        grdRecipe.TextMatrix(Row, 6) = rs.Fields("Unit_of_Measure") & ""
                        grdRecipe.TextMatrix(Row, 4) = 1
                        grdRecipe.TextMatrix(Row, 4) = Format(rs.Fields("Ave_Cost"), "0.00")
                        grdRecipe.TextMatrix(Row, 7) = Format(rs.Fields("Ave_Cost"), "0.00")
                        rs.Close
                    End If
            End Select
        Case 2
            If grdRecipe.Tag = grdRecipe.TextMatrix(Row, 2) Then Exit Sub
            Select Case grdRecipe.Tag
                Case "ml"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "lt"
                            grdRecipe.TextMatrix(Row, 3) = Val(grdRecipe.TextMatrix(Row, 3)) / 1000
                        Case "1 x 25ml"
                            grdRecipe.TextMatrix(Row, 3) = "25"
                         Case "2 x 50ml"
                            grdRecipe.TextMatrix(Row, 3) = "50"
                        Case "Whole Unit"
                            grdRecipe.TextMatrix(Row, 3) = grdRecipe.TextMatrix(Row, 5)
                    End Select
                Case "lt"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "ml"
                            grdRecipe.TextMatrix(Row, 3) = Val(grdRecipe.TextMatrix(Row, 3)) * 1000
                        Case "1 x 25ml"
                            grdRecipe.TextMatrix(Row, 3) = "0.025"
                         Case "2 x 50ml"
                            grdRecipe.TextMatrix(Row, 3) = "0.050"
                        Case "Whole Unit"
                            grdRecipe.TextMatrix(Row, 3) = grdRecipe.TextMatrix(Row, 5)
                    End Select
                Case "g"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "kg"
                            grdRecipe.TextMatrix(Row, 3) = Val(grdRecipe.TextMatrix(Row, 3)) / 1000
                    End Select
                Case "kg"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "g"
                            grdRecipe.TextMatrix(Row, 3) = Val(grdRecipe.TextMatrix(Row, 3)) * 1000
                    End Select
                Case "Single Tot"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "ml"
                            grdRecipe.TextMatrix(Row, 3) = "25"
                        Case "lt"
                            grdRecipe.TextMatrix(Row, 3) = "0.025"
                        Case "Whole Unit"
                            grdRecipe.TextMatrix(Row, 3) = grdRecipe.TextMatrix(Row, 5)
                    End Select
                Case "Double Tot"
                    Select Case grdRecipe.TextMatrix(Row, 2)
                        Case "ml"
                            grdRecipe.TextMatrix(Row, 3) = "50"
                        Case "lt"
                            grdRecipe.TextMatrix(Row, 3) = "0.050"
                        Case "Whole Unit"
                            grdRecipe.TextMatrix(Row, 3) = grdRecipe.TextMatrix(Row, 5)
                    End Select
            End Select
            If grdRecipe.TextMatrix(Row, 2) = "Single Tot" Then grdRecipe.TextMatrix(Row, 3) = "1 x 25ml"
            If grdRecipe.TextMatrix(Row, 2) = "Double Tot" Then grdRecipe.TextMatrix(Row, 3) = "2 x 25ml"
            If grdRecipe.TextMatrix(Row, 2) = "Whole Unit" Then grdRecipe.TextMatrix(Row, 3) = 1
        Case 3
            ActiveReadServer "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(Row, 1), InStrRev(grdRecipe.TextMatrix(Row, 1), ",") + 1) & "'"
            If rs.RecordCount > 0 Then
                grdRecipe.TextMatrix(Row, 5) = rs.Fields("Unit_Size") & ""
                If grdRecipe.TextMatrix(Row, 5) = 0 Then grdRecipe.TextMatrix(Row, 5) = "1"
                grdRecipe.TextMatrix(Row, 6) = rs.Fields("Unit_of_Measure") & ""
                grdRecipe.TextMatrix(Row, 4) = 1
                grdRecipe.TextMatrix(Row, 4) = Format(rs.Fields("Ave_Cost"), "0.00")
                grdRecipe.TextMatrix(Row, 7) = Format(rs.Fields("Ave_Cost"), "0.00")
            End If
            rs.Close
        Case 4
            grdRecipe.TextMatrix(Row, 4) = Format(grdRecipe.TextMatrix(Row, 4), "0.00")
    End Select
    If Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) > 0 Then
        Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
            Case "Stock Item", "Sales Item", "Stock Item (Hidden)", "Stock Item (Choice)", "Sales Item (Choice)"
                Select Case grdRecipe.TextMatrix(Row, 6)
                    Case "ml"
                        Select Case grdRecipe.TextMatrix(Row, 2)
                            Case "ml"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "Single Tot"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (25 / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "Double Tot"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (50 / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "Whole Unit"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)), "0.000")
                        End Select
                    Case "lt"
                        Select Case grdRecipe.TextMatrix(Row, 2)
                            Case "lt"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "ml"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * ((Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5)) / 1000)), "0.000")
                            Case "Single Tot"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * ((25 / Val(grdRecipe.TextMatrix(Row, 5)) / 1000)), "0.000")
                            Case "Double Tot"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * ((50 / Val(grdRecipe.TextMatrix(Row, 5)) / 1000)), "0.000")
                            Case "Whole Unit"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)), "0.000")
                        End Select
                    Case "g"
                        Select Case grdRecipe.TextMatrix(Row, 2)
                            Case "g"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                        End Select
                    Case "kg"
                        Select Case grdRecipe.TextMatrix(Row, 2)
                            Case "kg"
                                grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5))), "0.000")
                            Case "g"
                                If Val(grdRecipe.TextMatrix(Row, 5)) <> 0 Then
                                    grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * ((Val(grdRecipe.TextMatrix(Row, 3)) / Val(grdRecipe.TextMatrix(Row, 5)) / 1000)), "0.000")
                                Else
                                    grdRecipe.TextMatrix(Row, 4) = "0.00"
                                End If
                        End Select
                    Case "each"
                        grdRecipe.TextMatrix(Row, 4) = Format(Val(grdRecipe.TextMatrix(grdRecipe.Row, 7)) * (Val(grdRecipe.TextMatrix(Row, 3))), "0.000")
                End Select
        End Select
    End If
    totalcost = 0
    For i = 1 To grdRecipe.Rows - 1
        Select Case grdRecipe.TextMatrix(i, 0)
            Case "Stock Item", "Sales Item", "Stock Item (Hidden)"
                totalcost = totalcost + Val(grdRecipe.TextMatrix(i, 4))
        End Select
    Next i
    If totalcost > 0 Then
        txtLandCost.Text = Format(totalcost, "0.00")
    End If
    For i = 1 To grdRecipe.Rows - 1
        TotalProd = TotalProd + grdRecipe.ValueMatrix(i, 3)
    Next i
    If grdMain.ValueMatrix(grdMain.Row, 3) <> 0 Then
        If TotalProd <> 0 Then
            frmMain.Toolbar1.Buttons(2).Enabled = True
        Else
            frmMain.Toolbar1.Buttons(2).Enabled = False
        End If
    End If
End Sub
Private Sub grdRecipe_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    DoEvents
    grdMain.Tag = "1"
    grdRecipe.Tag = grdRecipe.Text
    cmdRefresh.Enabled = True
    grdMain.Tag = ""
End Sub
Private Sub grdRecipe_Click()
    If grdRecipe.Col = 4 Then
        totalcost = 0
        For i = 1 To grdRecipe.Rows - 1
            totalcost = totalcost + Val(grdRecipe.TextMatrix(grdRecipe.Row, 4))
        Next i
    End If
End Sub

Private Sub grdRecipe_EnterCell()
    totalcost = 0
    grdRecipe.TextMatrix(grdRecipe.Row, 5) = grdMain.TextMatrix(grdMain.Row, 0)
    For i = 1 To grdRecipe.Rows - 1
        If grdRecipe.TextMatrix(grdRecipe.Row, 5) = grdMain.TextMatrix(grdMain.Row, 0) Then
            totalcost = totalcost + Val(grdRecipe.TextMatrix(i, 4))
        End If
        
    Next i
    Select Case grdRecipe.Col
        Case 0
            grdRecipe.ColComboList(0) = "Stock Item|Message"
            If grdRecipe.Text = "" Then
                grdRecipe.Text = "Stock Item"
                grdRecipe.TextMatrix(grdRecipe.Row, 2) = " "
                grdRecipe.TextMatrix(grdRecipe.Row, 3) = " "
                grdRecipe.TextMatrix(grdRecipe.Row, 4) = " "
                grdRecipe.MergeRow(grdRecipe.Row) = True
                grdRecipe.Cell(flexcpBackColor, grdRecipe.Row, 2, grdRecipe.Row, 4) = &HE0E0E0
            End If
            grdRecipe.Editable = flexEDKbdMouse
            grdRecipe.ColComboList(1) = ""
        Case 1
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Message", "Price Size Change", "Exit"
                    grdRecipe.ColComboList(1) = ""
                    grdRecipe.Editable = flexEDKbdMouse
                    grdRecipe.ColComboList(2) = ""
                Case "Preparation Recipe"
                    If grdRecipe.ColComboList(1) = "" Then
                        Screen.MousePointer = 11
                        ActiveReadServer "Select * from Products where Unit_of_Measure ='Preparation Recipe' order by Description"
                        grdRecipe.ColComboList(1) = ""
                        While Not rs.EOF
                            grdRecipe.ColComboList(1) = grdRecipe.ColComboList(1) & "|" & rs.Fields("Description") & " ," & rs.Fields("Product_Code")
                            rs.MoveNext
                        Wend
                        rs.Close
                        Screen.MousePointer = 0
                    End If
                    grdRecipe.Editable = flexEDKbdMouse
                    grdRecipe.ColComboList(2) = ""
                Case "Stock Item", "Stock Item (Choice)"
                    If grdRecipe.ColComboList(1) = "" Then
                        Screen.MousePointer = 11
                        ActiveReadServer "Select Product_Code,Description,isnull(Unit_Size,'') as Unit_Size,Unit_of_Measure,Ave_Cost from Products where Stock_Item=1 and isnull(Production_Item,0) <> 1 and Product_Code in (Select Product_Code from Quantities where Location_No =" & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & ") order by Description "
                        While Not rs.EOF
                            If rs.Fields("Unit_Size") = "0" Then
                                UnitSize = ""
                            Else
                                UnitSize = rs.Fields("Unit_Size")
                            End If
                            grdRecipe.ColComboList(1) = grdRecipe.ColComboList(1) & "|" & rs.Fields("Description") & " " & UnitSize & rs.Fields("Unit_of_Measure") & " ," & rs.Fields("Product_Code")
                            rs.MoveNext
                        Wend
                        rs.Close
                        Screen.MousePointer = 0
                    End If
                    grdRecipe.Editable = flexEDKbdMouse
                Case "Stock Item (Hidden)"
                    If grdRecipe.ColComboList(1) = "" Then
                        Screen.MousePointer = 11
                        ActiveReadServer "Select Product_Code,Description,isnull(Unit_Size,'') as Unit_Size,Unit_of_Measure,Ave_Cost from Products where Stock_Item=1 order by Description"
                        While Not rs.EOF
                            If rs.Fields("Unit_Size") = "0" Then
                                UnitSize = ""
                            Else
                                UnitSize = rs.Fields("Unit_Size")
                            End If
                            grdRecipe.ColComboList(1) = grdRecipe.ColComboList(1) & "|" & rs.Fields("Description") & " " & UnitSize & rs.Fields("Unit_of_Measure") & " ," & rs.Fields("Product_Code")
                            rs.MoveNext
                        Wend
                        rs.Close
                        Screen.MousePointer = 0
                    End If
                    grdRecipe.Editable = flexEDKbdMouse
                Case "Sales Item", "Sales Item (Choice)"
                    If grdRecipe.ColComboList(1) = "" Then
                        Screen.MousePointer = 11
                        ActiveReadServer "Select Product_Code,Description,isnull(Unit_Size,'') as Unit_Size,Unit_of_Measure,Ave_Cost from Products where Sales_Item=1 order by Description"
                        While Not rs.EOF
                            If rs.Fields("Unit_Size") = "0" Then
                                UnitSize = ""
                            Else
                                UnitSize = rs.Fields("Unit_Size")
                            End If
                            grdRecipe.ColComboList(1) = grdRecipe.ColComboList(1) & "|" & rs.Fields("Description") & " " & UnitSize & rs.Fields("Unit_of_Measure") & " ," & rs.Fields("Product_Code")
                            rs.MoveNext
                        Wend
                        rs.Close
                        Screen.MousePointer = 0
                    End If
                    grdRecipe.Editable = flexEDKbdMouse
            End Select
        Case 2
            If grdRecipe.TextMatrix(grdRecipe.Row, 1) = "" Then
                grdRecipe.Col = 1
                Exit Sub
            End If
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Stock Item", "Sales Item", "Stock Item (Hidden)", "Stock Item (Choice)", "Sales Item (Choice)"
                    ActiveReadServer "Select Unit_of_Measure from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Row, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Row, 1), ",") + 1) & "'"
                    If rs.RecordCount > 0 Then
                        Select Case rs.Fields("Unit_of_Measure")
                            Case "ml"
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.ColComboList(2) = "ml|Single Tot|Double Tot|Whole Unit"
                                grdRecipe.Editable = flexEDKbdMouse
                            Case "lt"
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.ColComboList(2) = "ml|lt|Single Tot|Double Tot|Whole Unit"
                                grdRecipe.Editable = flexEDKbdMouse
                            Case "g"
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.ColComboList(2) = "g|Whole Unit"
                                grdRecipe.Editable = flexEDKbdMouse
                            Case "kg"
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.ColComboList(2) = "g|kg|Whole Unit"
                                grdRecipe.Editable = flexEDKbdMouse
                            Case Else
                                grdRecipe.ColComboList(2) = ""
                                grdRecipe.Editable = flexEDNone
                        End Select
                    End If
                    rs.Close
                Case Else
                    grdRecipe.Editable = flexEDNone
            End Select
        Case 3
            If grdRecipe.TextMatrix(grdRecipe.Row, 1) = "" Then
                grdRecipe.Col = 1
                Exit Sub
            End If
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Message", "Preparation Recipe", "Price/Size Change", "Exit"
                    grdRecipe.Editable = flexEDNone
                    grdRecipe.ColComboList(2) = ""
                Case Else
                    grdRecipe.Editable = flexEDKbdMouse
            End Select
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 2)
                Case "Whole Unit", "Single Tot", "Double Tot"
                    grdRecipe.Editable = flexEDNone
            End Select
        Case 4
            If grdRecipe.TextMatrix(grdRecipe.Row, 1) = "" Then
                grdRecipe.Col = 1
                Exit Sub
            End If
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Message", "Preparation Recipe", "Price/Size Change", "Exit"
                    grdRecipe.Editable = flexEDNone
                    grdRecipe.ColComboList(2) = ""
                Case "Sales Item (Choice)"
                    grdRecipe.Editable = flexEDKbdMouse
                Case Else
                    grdRecipe.Editable = flexEDNone
            End Select
    End Select
End Sub

Private Sub grdMain_GotFocus()
    grdMain.HighLight = flexHighlightWithFocus
    grdMain.SelectionMode = flexSelectionFree
End Sub

Private Sub grdRecipe_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 8, 48 To 57, 96 To 105
            If grdRecipe.Col = 3 Then grdRecipe.EditCell
        Case 46
            If grdRecipe.Row <> 0 Then
                grdRecipe.RemoveItem grdRecipe.Row
                If grdRecipe.Rows = 1 Then
                    grdRecipe.Rows = grdRecipe.Rows + 1
                End If
                totalcost = 0
                For i = 1 To grdRecipe.Rows - 1
                    totalcost = totalcost + Val(grdRecipe.TextMatrix(i, 4))
                Next i
                If totalcost > 0 Then
                    txtLandCost.Text = Format(totalcost, "0.00")
                End If
                grdMain.HighLight = flexHighlightAlways
                grdMain.SelectionMode = flexSelectionByRow
                grdMain.Enabled = False
                frmMain.Toolbar1.Buttons(2).Enabled = False
            End If
        Case 40
            If grdRecipe.Row = grdRecipe.Rows - 1 And grdRecipe.TextMatrix(grdRecipe.Row, 1) <> "" Then
                If grdRecipe.Rows < 46 Then
                    grdRecipe.Rows = grdRecipe.Rows + 1
                    grdRecipe.Col = 0
                End If
            Else
                If grdRecipe.TextMatrix(grdRecipe.Row, 0) = "Exit" Then
                    If grdRecipe.Rows < 46 Then
                        grdRecipe.Rows = grdRecipe.Rows + 1
                        grdRecipe.Col = 0
                    End If
                End If
            End If
        Case 38
            If grdRecipe.TextMatrix(grdRecipe.Row, 1) = "" And grdRecipe.Row = grdRecipe.Rows - 1 Then
                grdRecipe.RemoveItem grdRecipe.Row
            End If
    End Select
End Sub

Private Sub grdRecipe_KeyPress(KeyAscii As Integer)
    Select Case grdRecipe.Col
        Case 0
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
        Case 1
            Select Case grdRecipe.TextMatrix(grdRecipe.Row, 0)
                Case "Stock Item", "Sales Item", "Preparation Recipe", "Stock Item (Hidden)"
                    Select Case KeyAscii
                    Case 8, 13, 27
                    Case Else
                        KeyAscii = 0
                End Select
            End Select
        Case 2
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
        Case 3
            If grdRecipe.TextMatrix(grdRecipe.Row, 0) = "Stock Item" Or grdRecipe.TextMatrix(grdRecipe.Row, 0) = "Stock Item (Hidden)" Then
                Select Case KeyAscii
                    Case 8, 13, 27
                    Case 46, 48 To 57
                    Case Else
                        KeyAscii = 0
                End Select
            Else
                Select Case KeyAscii
                    Case 8, 13, 27
                    Case 48 To 57
                    Case Else
                        KeyAscii = 0
                End Select
            End If
        Case 4
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
    End Select
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub grdRecipe_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case grdRecipe.Col
        Case 0
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
        Case 2
            Select Case KeyAscii
                Case 8, 13, 27
                Case Else
                    KeyAscii = 0
            End Select
        Case 3
            If grdRecipe.TextMatrix(Row, 0) = "Stock Item" Or grdRecipe.TextMatrix(Row, 0) = "Stock Item (Hidden)" Then
                Select Case KeyAscii
                    Case 8, 13, 27
                    Case 46, 48 To 57
                    Case Else
                        KeyAscii = 0
                End Select
            Else
                Select Case KeyAscii
                    Case 8, 13, 27
                    Case 48 To 57
                    Case Else
                        KeyAscii = 0
                End Select
            End If
            DoEvents
            If InStr(grdRecipe.EditText, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
        Case 4
            Select Case KeyAscii
                Case 48 To 57
                Case 8, 13, 27, 45, 46
                Case Else
                    KeyAscii = 0
            End Select
            If InStr(grdRecipe.EditText, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    End Select
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub grdMain_LostFocus()
    grdMain.HighLight = flexHighlightAlways
    grdMain.SelectionMode = flexSelectionByRow
End Sub

