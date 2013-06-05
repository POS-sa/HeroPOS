VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmProdSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Product Code Lookup"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSupp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      TabIndex        =   2
      Top             =   420
      Width           =   2475
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   285
      Left            =   3690
      TabIndex        =   1
      Top             =   90
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
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
   Begin VB.TextBox txtProd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Top             =   90
      Width           =   2475
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   285
      Left            =   3690
      TabIndex        =   5
      Top             =   420
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
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
   Begin VB.Label Label1 
      Caption         =   "Product Code:"
      Height          =   195
      Left            =   70
      TabIndex        =   4
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "Supplier Code:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   450
      Width           =   1305
   End
End
Attribute VB_Name = "frmProdSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonEx1_Click()
    ActiveReadServer "Select Product_Code,  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
    "+ Unit_of_Measure END AS Description from Products where Stock_Item = 1 and Department_No in" & _
    " (Select Department_No from Departments_Stock where Location_No = " & Trim(Mid(frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 2), 1, InStr(frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 2), "-") - 1)) & ") and Product_Code = '" & txtProd.Text & "'"
    If rs.RecordCount > 0 Then
        frmGRV.Tag = rs.Fields("Description") & " - " & rs.Fields("Product_Code")
    Else
        frmGRV.Tag = ""
    End If
    rs.Close
    Unload Me
End Sub
Private Sub ButtonEx2_Click()
    ActiveReadServer "Select Product_Code from Supplier_Links where Supplier_Code = '" & txtSupp.Text & "' and Supplier_no = '" & Replace(Mid(frmGRV.cmbSuppliers, InStrRev(frmGRV.cmbSuppliers, "(") + 1), ")", "") & "'"
    If rs.RecordCount > 0 Then
        txtProd.Text = rs.Fields("Product_Code")
    Else
        txtSupp.Text = ""
    End If
    rs.Close
    ActiveReadServer "Select Product_Code,  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
    "+ Unit_of_Measure END AS Description from Products where Stock_Item = 1 and Department_No in" & _
    " (Select Department_No from Departments_Stock where Location_No = " & Trim(Mid(frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 2), 1, InStr(frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 2), "-") - 1)) & ") and Product_Code = '" & txtProd.Text & "'"
    If rs.RecordCount > 0 Then
        frmGRV.Tag = rs.Fields("Description") & " - " & rs.Fields("Product_Code")
    Else
        frmGRV.Tag = ""
    End If
    rs.Close
    Unload Me
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
        ActiveReadServer "Select Product_Code,  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
        "+ Unit_of_Measure END AS Description from Products where Stock_Item = 1 and Department_No in" & _
        " (Select Department_No from Departments_Stock where Location_No = " & Trim(Mid(frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 2), 1, InStr(frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 2), "-") - 1)) & ") and Product_Code = '" & txtProd.Text & "'"
        If rs.RecordCount > 0 Then
            frmGRV.Tag = rs.Fields("Description") & " - " & rs.Fields("Product_Code")
        Else
            frmGRV.Tag = ""
        End If
        rs.Close
        Unload Me
    End If
    If KeyCode = 40 Or KeyCode = 38 Then
        KeyCode = 0
        txtSupp.SetFocus
    End If
End Sub
Private Sub txtSupp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ActiveReadServer "Select Product_Code from Supplier_Links where Supplier_Code = '" & txtSupp.Text & "' and Supplier_no = '" & Replace(Mid(frmGRV.cmbSuppliers, InStrRev(frmGRV.cmbSuppliers, "(") + 1), ")", "") & "'"
        If rs.RecordCount > 0 Then
            txtProd.Text = rs.Fields("Product_Code")
        Else
            txtSupp.Text = ""
        End If
        rs.Close
        ActiveReadServer "Select Product_Code,  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
        "+ Unit_of_Measure END AS Description from Products where Stock_Item = 1 and Department_No in" & _
        " (Select Department_No from Departments_Stock where Location_No = " & Trim(Mid(frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 2), 1, InStr(frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 2), "-") - 1)) & ") and Product_Code = '" & txtProd.Text & "'"
        If rs.RecordCount > 0 Then
            frmGRV.Tag = rs.Fields("Description") & " - " & rs.Fields("Product_Code")
        Else
            frmGRV.Tag = ""
        End If
        rs.Close
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
