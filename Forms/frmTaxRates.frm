VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTaxRates 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tax Rates"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   0
      Left            =   2370
      TabIndex        =   2
      Top             =   3480
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
      Left            =   3510
      TabIndex        =   3
      Top             =   3480
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
      Left            =   1200
      TabIndex        =   1
      Top             =   3480
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
      Height          =   3330
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4545
      _cx             =   8017
      _cy             =   5874
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
      BackColorSel    =   15329975
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
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
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTaxRates.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
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
Attribute VB_Name = "frmTaxRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForms_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
        Case 2
            ActiveUpdateServer "Delete from Tax_Rates"
            frmDetails.lblTax.Caption = ""
            For i = 1 To grdGrid.Rows - 1
                If Trim(grdGrid.TextMatrix(i, 1)) <> "" Then
                    ActiveUpdateServer "INSERT INTO [Tax_Rates]([Tax_Type],[Tax_Rate], [Description]) VALUES (" & grdGrid.TextMatrix(i, 0) & "," & Val(Left(grdGrid.TextMatrix(i, 1), Len(grdGrid.TextMatrix(i, 1)) - 1)) & ",'" & Trim(grdGrid.TextMatrix(i, 2)) & "')"
                End If
                If grdGrid.TextMatrix(i, 3) = "-1" Then
                    If i = grdGrid.Rows - 1 Then
                        frmDetails.lblTax.Caption = frmDetails.lblTax.Caption & grdGrid.TextMatrix(i, 0)
                    Else
                        frmDetails.lblTax.Caption = frmDetails.lblTax.Caption & grdGrid.TextMatrix(i, 0) & "-"
                    End If
                End If
            Next i
            If frmDetails.lblTax.Caption = "" Then frmDetails.lblTax.Caption = "1"
            If Right(frmDetails.lblTax.Caption, 1) = "-" Then frmDetails.lblTax.Caption = Trim(Mid(frmDetails.lblTax.Caption, 1, Len(frmDetails.lblTax.Caption) - 1))
            MsgBox "Taxrate Updated.", vbInformation, "HeroPOS Information Message"
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
    If grdGrid.Rows > 2 Then grdGrid.Row = 1
End Sub
Private Sub Form_Load()
    Dim ratesplit As Variant
    i = 0
    ratesplit = Split(frmDetails.lblTax.Caption, "-", -1)
    ActiveReadServer "Select * from Tax_Rates order by Tax_Type"
    If rs.RecordCount > 0 Then
        grdGrid.Rows = rs.RecordCount + 1
    End If
    While Not rs.EOF
        i = i + 1
        grdGrid.TextMatrix(i, 0) = rs.Fields("Tax_Type")
        grdGrid.TextMatrix(i, 1) = Format(rs.Fields("Tax_Rate"), "0.00") & "%"
        grdGrid.TextMatrix(i, 2) = rs.Fields("Description")
        For b = 0 To UBound(ratesplit)
            If rs.Fields("Tax_Type") = Val(ratesplit(b)) Then grdGrid.TextMatrix(i, 3) = "-1"
        Next b
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub grdGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        If Right(grdGrid.Text, 1) <> "%" And Trim(grdGrid.Text) <> "" Then
            grdGrid.Text = Format(grdGrid.Text, "0.00") & "%"
        End If
    End If
End Sub
Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
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
                    grdGrid.TextMatrix(grdGrid.Row, 0) = Val(grdGrid.TextMatrix(grdGrid.Row - 1, 0)) + 1
                End If
            End If
    End Select
End Sub
Private Sub grdGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = 0
    End Select
    If Col = 1 Then
        Select Case KeyAscii
        Case 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
    End If
End Sub

