VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSuppCodes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Supplier Codes..."
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   0
      Left            =   4050
      TabIndex        =   2
      Top             =   3150
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   1
      Left            =   5190
      TabIndex        =   3
      Top             =   3150
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Help"
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
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   2
      Left            =   2880
      TabIndex        =   1
      Top             =   3150
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Ok"
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
   Begin VSFlex8Ctl.VSFlexGrid grdGrid 
      Height          =   3030
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6255
      _cx             =   11033
      _cy             =   5345
      Appearance      =   1
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSuppCodes.frx":0000
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
      ExplorerBar     =   0
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
End
Attribute VB_Name = "frmSuppCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForms_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Unload Me
        Case 2
            If frmSuppCodes.Tag = "Supp" Then
                ActiveUpdateServer "Delete from Supplier_Links where Product_Code = '" & frmSuppliers.grdSupp.TextMatrix(frmSuppliers.grdSupp.Row, 0) & "'"
                For i = 1 To grdGrid.Rows - 1
                    If Trim(grdGrid.TextMatrix(i, 1)) <> "" Then
                        ActiveUpdateServer "INSERT INTO [Supplier_Links]([Supplier_No], [Product_Code],[Supplier_Code],List_Price,Date_Time) VALUES ('" & Trim(Mid(grdGrid.TextMatrix(i, 0), InStr(grdGrid.TextMatrix(i, 0), "-") + 1)) & "','" & frmSuppliers.grdSupp.TextMatrix(frmSuppliers.grdSupp.Row, 0) & "','" & Trim(grdGrid.TextMatrix(i, 1)) & "','" & grdGrid.ValueMatrix(i, 2) & "',Getdate())"
                    End If
                Next i
            Else
                ActiveUpdateServer "Delete from Supplier_Links where Product_Code = '" & frmProducts.txtProductCode.Text & "'"
                For i = 1 To grdGrid.Rows - 1
                    If Trim(grdGrid.TextMatrix(i, 0)) <> "" Then
                        ActiveUpdateServer "INSERT INTO [Supplier_Links]([Supplier_No], [Product_Code],[Supplier_Code],List_Price,Date_Time) VALUES ('" & Trim(Mid(grdGrid.TextMatrix(i, 0), InStr(grdGrid.TextMatrix(i, 0), "-") + 1)) & "','" & frmProducts.txtProductCode.Text & "','" & Trim(grdGrid.TextMatrix(i, 1)) & "','" & grdGrid.ValueMatrix(i, 2) & "',Getdate())"
                    End If
                Next i
            End If
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 11
    DoEvents
    If grdGrid.Rows > 1 Then
        grdGrid.Row = 1
        grdGrid.Col = 0
    End If
    grdGrid.Rows = 1
    If frmSuppCodes.Tag = "Supp" Then
        ActiveReadServer2 "Select Supplier_No,Supplier_Code,(Select Supplier_Name from Suppliers where Supplier_Links.Supplier_No = Suppliers.Supplier_No) as Supplier_Name,Product_Code,List_Price from Supplier_Links where Product_code = '" & frmSuppliers.grdSupp.TextMatrix(frmSuppliers.grdSupp.Row, 0) & "' order by Supplier_Name"
    Else
        ActiveReadServer2 "Select Supplier_No,Supplier_Code,(Select Supplier_Name from Suppliers where Supplier_Links.Supplier_No = Suppliers.Supplier_No) as Supplier_Name,Product_Code,List_Price from Supplier_Links where Product_code = '" & frmProducts.txtProductCode & "' order by Supplier_Name"
    End If
    While Not rs2.EOF
        grdGrid.Rows = grdGrid.Rows + 1
        grdGrid.Row = grdGrid.Rows - 1
        grdGrid.TextMatrix(grdGrid.Rows - 1, 0) = rs2.Fields("Supplier_Name") & " - " & rs2.Fields("Supplier_No")
        grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = rs2.Fields("Supplier_Code")
        grdGrid.TextMatrix(grdGrid.Rows - 1, 2) = Format(rs2.Fields("List_Price"), "0.00")
        rs2.MoveNext
    Wend
    rs2.Close
    If grdGrid.Rows = 1 Then grdGrid.Rows = 2
    Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
    grdGrid.TextMatrix(0, 0) = "Supplier"
    grdGrid.TextMatrix(0, 1) = "Supplier_Code"
    grdGrid.TextMatrix(0, 2) = "List Price"
    grdGrid.ColWidth(0) = grdGrid.Width * 0.5
    grdGrid.ColWidth(1) = grdGrid.Width * 0.3
    grdGrid.ColWidth(2) = grdGrid.Width * 0.2
    grdGrid.ColAlignment(0) = flexAlignLeftCenter
    grdGrid.ColAlignment(1) = flexAlignLeftCenter
    grdGrid.ColAlignment(2) = flexAlignRightCenter
End Sub
Private Sub grdGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If grdGrid.Col = 2 Then grdGrid.TextMatrix(Row, Col) = Format(grdGrid.TextMatrix(Row, Col), "0.00")
End Sub
Private Sub grdGrid_EnterCell()
    Select Case Col
        Case 0
            grdGrid.ColComboList(0) = ""
            ActiveReadServer "Select Supplier_No, Supplier_Name from Suppliers order by Supplier_Name desc"
            While Not rs.EOF
                grdGrid.ColComboList(0) = rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No") & "|" & grdGrid.ColComboList(0)
                rs.MoveNext
            Wend
            rs.Close
            grdGrid.Editable = flexEDNone
        Case 1, 2
            If grdGrid.TextMatrix(grdGrid.Row, grdGrid.Col) = "" Then
                grdGrid.Editable = flexEDNone
            Else
                grdGrid.Editable = flexEDKbdMouse
            End If
    End Select
End Sub
Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            grdGrid.Editable = flexEDKbdMouse
        Case 46
            grdGrid.RemoveItem grdGrid.Row
        Case 38
            If grdGrid.Row = grdGrid.Rows - 1 Then
                If grdGrid.Text = "" Then
                    grdGrid.RemoveItem grdGrid.Rows - 1
                End If
            End If
        Case 40
            If grdGrid.Row = grdGrid.Rows - 1 Then
                If grdGrid.Text <> "" Then
                    grdGrid.Rows = grdGrid.Rows + 1
                    grdGrid.Row = grdGrid.Rows - 1
                    grdGrid.Col = 0
                End If
            End If
        Case 45, 48 To 57, 65 To 90, 96 To 105, 109, 110, 189
            grdGrid.EditCell
    End Select
End Sub
Private Sub grdGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case 1
            Select Case KeyAscii
        Case 8, 45, 46, 48 To 57
        Case 39
            KeyAscii = 0
        Case 65 To 90
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case Else
            KeyAscii = 0
    End Select
        Case 2
            If InStr(grdGrid.EditText, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
            Select Case KeyAscii
                Case 8, 46, 48 To 57
                Case Else: KeyAscii = 0
            End Select
    End Select
End Sub


