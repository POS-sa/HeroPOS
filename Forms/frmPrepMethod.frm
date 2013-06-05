VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPrepMethod 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Preparation Method"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProdCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   210
      Width           =   2265
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   6435
   End
   Begin btButtonEx.ButtonEx cmdRates 
      Height          =   345
      Left            =   6930
      TabIndex        =   1
      Top             =   6840
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Appearance      =   3
      AutoMask        =   0   'False
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
   End
   Begin btButtonEx.ButtonEx cmdOk 
      Height          =   345
      Left            =   5580
      TabIndex        =   2
      Top             =   6840
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Appearance      =   3
      AutoMask        =   0   'False
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
   End
   Begin VSFlex8Ctl.VSFlexGrid grdRecipe 
      Height          =   3120
      Left            =   60
      TabIndex        =   7
      Top             =   1020
      Width           =   8140
      _cx             =   14358
      _cy             =   5503
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
      SheetBorder     =   -2147483632
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPrepMethod.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   2
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
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
      ComboSearch     =   1
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
   Begin RichTextLib.RichTextBox txtRemarks 
      Height          =   2355
      Left            =   225
      TabIndex        =   0
      ToolTipText     =   " Preparation Method "
      Top             =   4320
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   4154
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmPrepMethod.frx":0078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   345
      Left            =   60
      TabIndex        =   8
      Top             =   6870
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Print Recipe"
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
   Begin MSForms.Image Image2 
      Height          =   3165
      Left            =   60
      Top             =   1020
      Width           =   8175
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "14420;5583"
   End
   Begin MSForms.Image Image1 
      Height          =   2565
      Index           =   8
      Left            =   60
      Top             =   4230
      Width           =   8175
      BackColor       =   16777215
      BorderStyle     =   0
      MousePointer    =   8
      SpecialEffect   =   3
      Size            =   "14420;4524"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code: "
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   210
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description: "
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   570
      Width           =   1335
   End
   Begin MSForms.Image Image11 
      Height          =   345
      Left            =   1530
      Top             =   150
      Width           =   2385
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "4207;609"
   End
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   1530
      Top             =   540
      Width           =   6585
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "11615;609"
   End
   Begin MSForms.Image Image4 
      Height          =   915
      Left            =   60
      Top             =   60
      Width           =   8145
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "14367;1614"
   End
End
Attribute VB_Name = "frmPrepMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
On Error GoTo trap
With frmSlipDetails
    PrintErr = 0
    Slip_Port = ""
    
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As #filenum
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
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
    End If

    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "PRODUCT RECIPE"
    Print #filenum, txtDescription.Text
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    ActiveReadServer2 "Select * from Recipes where Product_Code = '" & txtProdCode.Text & "'"  '
    While Not rs2.EOF
        Print #filenum, Trim(Mid(rs2.Fields("Description"), InStrRev(rs2.Fields("Description"), ",") + 1)) & " - " & Trim(Mid(rs2.Fields("Description"), 1, InStrRev(rs2.Fields("Description"), ",") - 1))
        Print #filenum, "Using " & rs2.Fields("Qty_Used") & " (" & rs2.Fields("Unit_of_Measure") & ")"
        rs2.MoveNext
        If rs2.EOF = False Then
            If Slip_Printer_Type = 0 Then
                Print #filenum, Chr(27) & Chr(77) & Chr(49);
                Print #filenum, String(40, "-")
            Else
                Print #filenum, String(33, "-")
            End If
        End If
    Wend
    If rs2.RecordCount = 0 Then
        Print #filenum, "NO RECIPE LOADED"
    End If
    rs2.Close
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "-")
    Else
        Print #filenum, String(33, "-")
    End If
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "PREPARATION METHOD"
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, txtRemarks.Text
    For i = 1 To .grdFoot.Rows - 1
        If Trim(.grdFoot.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdFoot.TextMatrix(i, 1))
                Case "Line Feeds"
                    Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdFoot.TextMatrix(i, 0)));
                Case Else
                    Select Case .grdFoot.TextMatrix(i, 2)
                        Case "Left": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdFoot.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdFoot.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #filenum
End With
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
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    frmError.Show vbModal
    Close #filenum
    On Error GoTo 0
End Sub
Private Sub cmdOk_Click()
    ActiveUpdateServer "Delete from Preparations  where Product_code ='" & txtProdCode.Text & "'"
    DoEvents
    ActiveUpdateServer "INSERT INTO [Preparations]([Product_Code], [Prep_Method])" & _
    " VALUES('" & txtProdCode.Text & "','" & txtRemarks.Text & "')"
    Unload Me
End Sub
Private Sub cmdRates_Click()
    Unload Me
End Sub

Private Sub Form_Load()
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
    txtProdCode.Text = frmProducts.txtProductCode
    txtDescription.Text = frmProducts.txtDescription.Text
    ActiveReadServer "Select * from Recipes where Product_Code= '" & txtProdCode.Text & "' order by Line_No"
    While Not rs.EOF
        grdRecipe.Rows = grdRecipe.Rows + 1
        grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1) = rs.Fields("Description")
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
                ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                If rs1.RecordCount > 0 Then
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                End If
                rs1.Close
            Case 3
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                If rs1.RecordCount > 0 Then
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                End If
                rs1.Close
            Case 4
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item (Hidden)"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                If rs1.RecordCount > 0 Then
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                End If
                rs1.Close
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
                ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                If rs1.RecordCount > 0 Then
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                End If
                rs1.Close
            Case 7
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 0) = "Stock Item (Choice)"
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 2) = rs.Fields("Unit_of_Measure")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 3) = rs.Fields("Qty_Used")
                grdRecipe.TextMatrix(grdRecipe.Rows - 1, 4) = Format(rs.Fields("Cost"), "0.000")
                ActiveReadServer1 "Select Unit_Size,Unit_of_Measure,Ave_Cost from products where product_Code='" & Mid(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), InStrRev(grdRecipe.TextMatrix(grdRecipe.Rows - 1, 1), ",") + 1) & "'"
                If rs1.RecordCount > 0 Then
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 5) = rs1.Fields("Unit_Size") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 6) = rs1.Fields("Unit_of_Measure") & ""
                    grdRecipe.TextMatrix(grdRecipe.Rows - 1, 7) = Format(rs1.Fields("Ave_Cost"), "0.00")
                End If
                rs1.Close
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
    ActiveReadServer "Select * from Preparations where Product_Code ='" & txtProdCode.Text & "'"
    If rs.RecordCount > 0 Then
        txtRemarks.Text = rs.Fields("Prep_Method") & ""
    End If
    rs.Close
End Sub
