VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmProdChange 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Product Code Changer..."
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBC 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   3330
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   450
      Visible         =   0   'False
      Width           =   345
      Begin VB.Label lblCode 
         BackColor       =   &H00C0FFC0&
         Caption         =   "BC"
         Height          =   165
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   30
         Width           =   285
      End
   End
   Begin VB.PictureBox PicBC 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   3330
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   345
      Begin VB.Label lblCode 
         BackColor       =   &H00C0FFC0&
         Caption         =   "BC"
         Height          =   165
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   30
         Width           =   285
      End
   End
   Begin btButtonEx.ButtonEx cmdBarcode 
      Height          =   630
      Left            =   3750
      TabIndex        =   5
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1111
      Appearance      =   3
      Caption         =   "Generate Barcode"
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
   Begin VB.TextBox txtProd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2715
   End
   Begin VB.TextBox txtSupp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   450
      Width           =   2715
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   630
      Left            =   4860
      TabIndex        =   1
      Top             =   120
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1111
      Appearance      =   3
      Caption         =   "GO>"
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "New Code:"
      Height          =   195
      Left            =   -405
      TabIndex        =   4
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Code:"
      Height          =   195
      Left            =   -405
      TabIndex        =   3
      Top             =   150
      Width           =   1305
   End
End
Attribute VB_Name = "frmProdChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
    If txtProd.Text = "" Or txtSupp.Text = "" Then
        MsgBox "Blank Codes are not Allowed", vbCritical, "HeroPOS"
        Exit Sub
    End If
    ActiveUpdateServer "ALTER TABLE Product_Prices NOCHECK CONSTRAINT FK_Product_Prices_Products"
    ActiveUpdateServer " Update Products set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
    ActiveUpdateServer "ALTER TABLE Product_Prices CHECK CONSTRAINT FK_Product_Prices_Products"
    DoEvents
    ActiveUpdateServer " Update Product_Prices set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
    DoEvents
    ActiveUpdateServer " Update Supplier_Links set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
    DoEvents
    ActiveUpdateServer " Update Recipes set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
    DoEvents
    ActiveUpdateServer " Update Recipes set Line_Code = '" & txtSupp.Text & "' where Line_Code = '" & txtProd.Text & "'"
    DoEvents
    ActiveUpdateServer " Update Preparations set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
    DoEvents
    Select Case frmProdChange.Tag
        Case "Products"
            frmProducts.txtProductCode = txtSupp.Text
            frmProducts.grdProd.TextMatrix(frmProducts.grdProd.Row, 0) = txtSupp.Text
        Case "GRV"
            frmProd.txtProductCode = txtSupp.Text
            frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0) = txtSupp.Text
    End Select
    Unload Me
End Sub
Private Sub cmdBarcode_Click()
    Randomize
    random = Trim(Str(Int(Rnd * 1000000) + 1))
    newcode = Trim(CodePrefix & String(10 - Len(random), "0")) & Trim(random)
    TopRow = 0
    BottomRow = 0
    For i = 1 To Len(newcode)
        If Len(newcode) < 9 Then
            If i / 2 <> Int(i / 2) Then
                TopRow = TopRow + Val(Mid(newcode, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(newcode, i, 1))
            End If
        Else
            If i / 2 = Int(i / 2) Then
                TopRow = TopRow + Val(Mid(newcode, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(newcode, i, 1))
            End If
        End If
    Next i
    TopRow = TopRow * 3
    Result = TopRow + BottomRow
    Result = Round((1 - ((Result / 10) - Int((Result / 10)))) * 10, 0)
    If Result = 10 Then Result = 0
    newcode = newcode & Trim(Result)
    PicBC(1).Visible = True
    txtSupp.Text = newcode
End Sub
Private Sub Form_Activate()
    Select Case frmProdChange.Tag
        Case "Products"
            txtProd.Text = frmProducts.txtProductCode
        Case "GRV"
            txtProd.Text = frmProd.txtProductCode
            frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0) = txtSupp.Text
    End Select
End Sub
Private Sub txtProd_Change()
    TopRow = 0
    BottomRow = 0
    For i = 1 To Len(txtProd.Text) - 1
        If Len(txtProd.Text) < 9 Then
            If i / 2 <> Int(i / 2) Then
                TopRow = TopRow + Val(Mid(txtProd.Text, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(txtProd.Text, i, 1))
            End If
        Else
            If i / 2 = Int(i / 2) Then
                TopRow = TopRow + Val(Mid(txtProd.Text, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(txtProd.Text, i, 1))
            End If
        End If
    Next i
    TopRow = TopRow * 3
    Result = TopRow + BottomRow
    Result = (1 - ((Result / 10) - Int((Result / 10)))) * 10
    If Result = 10 Then Result = 0
    If Round(Result, 0) = Int(Val(Right(txtProd, 1))) Then
        PicBC(0).Visible = True
    Else
        PicBC(0).Visible = False
    End If
End Sub

Private Sub txtProd_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 32
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub
Private Sub txtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtProd.Text = "" Or txtSupp.Text = "" Then
            MsgBox "Blank Codes are not Allowed", vbCritical, "HeroPOS"
            Exit Sub
        End If
        ActiveUpdateServer "ALTER TABLE Product_Prices NOCHECK CONSTRAINT FK_Product_Prices_Products"
        ActiveReadServer " Update Products set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
        ActiveUpdateServer "ALTER TABLE Product_Prices CHECK CONSTRAINT FK_Product_Prices_Products"
        DoEvents
        ActiveReadServer " Update Product_Prices set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Supplier_Links set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Recipes set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Recipes set Line_Code = '" & txtSupp.Text & "' where Line_Code = '" & txtProd.Text & "'"
        DoEvents
        Select Case frmProdChange.Tag
            Case "Products"
                frmProducts.txtProductCode = txtSupp.Text
                frmProducts.grdProd.TextMatrix(frmProducts.grdProd.Row, 0) = txtSupp.Text
            Case "GRV"
                frmProd.txtProductCode = txtSupp.Text
                frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0) = txtSupp.Text
        End Select
        Unload Me
    End If
    If KeyCode = 40 Or KeyCode = 38 Then
        KeyCode = 0
        txtSupp.SetFocus
    End If
End Sub
Private Sub txtSupp_Change()
    TopRow = 0
    BottomRow = 0
    For i = 1 To Len(txtSupp.Text) - 1
        If Len(txtSupp.Text) < 9 Then
            If i / 2 <> Int(i / 2) Then
                TopRow = TopRow + Val(Mid(txtSupp.Text, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(txtSupp.Text, i, 1))
            End If
        Else
            If i / 2 = Int(i / 2) Then
                TopRow = TopRow + Val(Mid(txtSupp.Text, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(txtSupp.Text, i, 1))
            End If
        End If
    Next i
    TopRow = TopRow * 3
    Result = TopRow + BottomRow
    Result = (1 - ((Result / 10) - Int((Result / 10)))) * 10
    If Result = 10 Then Result = 0
    If Round(Result, 0) = Int(Val(Right(txtSupp, 1))) Then
        PicBC(1).Visible = True
    Else
        PicBC(1).Visible = False
    End If
End Sub

Private Sub txtSupp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        If txtProd.Text = "" Or txtSupp.Text = "" Then
            MsgBox "Blank Codes are not Allowed", vbCritical, "HeroPOS"
            Exit Sub
        End If
        ActiveUpdateServer "ALTER TABLE Product_Prices NOCHECK CONSTRAINT FK_Product_Prices_Products"
        ActiveReadServer " Update Products set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
        ActiveUpdateServer "ALTER TABLE Product_Prices CHECK CONSTRAINT FK_Product_Prices_Products"
        DoEvents
        ActiveReadServer " Update Product_Prices set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Supplier_Links set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Recipes set Product_Code = '" & txtSupp.Text & "' where Product_Code = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Recipes set Line_Code = '" & txtSupp.Text & "' where Line_Code = '" & txtProd.Text & "'"
        DoEvents
        Select Case frmProdChange.Tag
            Case "Products"
                frmProducts.txtProductCode = txtSupp.Text
                frmProducts.grdProd.TextMatrix(frmProducts.grdProd.Row, 0) = txtSupp.Text
            Case "GRV"
                frmProd.txtProductCode = txtSupp.Text
                frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0) = txtSupp.Text
        End Select
        Unload Me
    End If
    If KeyCode = 40 Or KeyCode = 38 Then
        KeyCode = 0
        txtProd.SetFocus
    End If
End Sub
Private Sub txtSupp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 45, 46
        Case 39
            KeyAscii = Asc("`")
        Case 32
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub



