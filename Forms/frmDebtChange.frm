VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDebtChange 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Debtor Number Changer..."
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1110
      TabIndex        =   2
      Top             =   120
      Width           =   2715
   End
   Begin VB.TextBox txtSupp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1110
      TabIndex        =   0
      Top             =   450
      Width           =   2715
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   630
      Left            =   3870
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
      Caption         =   "New Number:"
      Height          =   195
      Left            =   -255
      TabIndex        =   4
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Number:"
      Height          =   195
      Left            =   -255
      TabIndex        =   3
      Top             =   150
      Width           =   1305
   End
End
Attribute VB_Name = "frmDebtChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
    If txtProd.Text = "" Or txtSupp.Text = "" Then
        MsgBox "Blank Numbers are not Allowed", vbCritical, "HeroPOS"
        Exit Sub
    End If
    ActiveReadServer " Select * from Debtors where Debtor_No = '" & txtSupp.Text & "'"
    If rs.RecordCount > 0 Then
        MsgBox "You Allready have a Debtor with this Number.", vbCritical, "HeroPOS"
        txtSupp.Text = ""
        txtSupp.SetFocus
        Exit Sub
    End If
    ActiveUpdateServer " Update Debtors set Debtor_No = '" & txtSupp.Text & "' where Debtor_No = '" & txtProd.Text & "'"
    DoEvents
    ActiveReadServer " Update Sales_Journal set Account_No = '" & txtSupp.Text & "' where Account_No = '" & txtProd.Text & "'"
    DoEvents
    ActiveReadServer " Update Debtor_Accounts set Account_No = '" & txtSupp.Text & "' where Account_No = '" & txtProd.Text & "'"
    DoEvents
    ActiveReadServer " Update Debtor_Discounts set Debtor_No = '" & txtSupp.Text & "' where Debtor_No = '" & txtProd.Text & "'"
    DoEvents
    frmGuests.txtSuppNo.Text = txtSupp.Text
    frmGuests.grdSupp.TextMatrix(frmGuests.grdSupp.Row, 0) = txtSupp.Text
    Unload Me
End Sub
Private Sub Form_Activate()
    txtProd.Text = frmGuests.txtSuppNo
End Sub
Private Sub txtProd_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 32
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtProd.Text = "" Or txtSupp.Text = "" Then
            MsgBox "Blank Numbers are not Allowed", vbCritical, "HeroPOS"
            Exit Sub
        End If
        ActiveReadServer " Select * from Debtors where Debtor_No = '" & txtSupp.Text & "'"
        If rs.RecordCount > 0 Then
            MsgBox "You Allready have a Debtor with this Number.", vbCritical, "HeroPOS"
            txtSupp.Text = ""
            txtSupp.SetFocus
            Exit Sub
        End If
        ActiveUpdateServer " Update Debtors set Debtor_No = '" & txtSupp.Text & "' where Debtor_No = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Sales_Journal set Account_No = '" & txtSupp.Text & "' where Account_No = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Debtor_Accounts set Account_No = '" & txtSupp.Text & "' where Account_No = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Debtor_Discounts set Debtor_No = '" & txtSupp.Text & "' where Debtor_No = '" & txtProd.Text & "'"
        DoEvents
        frmGuests.txtSuppNo.Text = txtSupp.Text
        frmGuests.grdSupp.TextMatrix(frmGuests.grdSupp.Row, 0) = txtSupp.Text
        Unload Me
    End If
    If KeyCode = 40 Or KeyCode = 38 Then
        KeyCode = 0
        txtSupp.SetFocus
    End If
End Sub
Private Sub txtSupp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtProd.Text = "" Or txtSupp.Text = "" Then
            MsgBox "Blank Numbers are not Allowed", vbCritical, "HeroPOS"
            Exit Sub
        End If
        ActiveReadServer " Select * from Debtors where Debtor_No = '" & txtSupp.Text & "'"
        If rs.RecordCount > 0 Then
            MsgBox "You Allready have a Debtor with this Number.", vbCritical, "HeroPOS"
            txtSupp.Text = ""
            txtSupp.SetFocus
            Exit Sub
        End If
        ActiveUpdateServer " Update Debtors set Debtor_No = '" & txtSupp.Text & "' where Debtor_No = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Sales_Journal set Account_No = '" & txtSupp.Text & "' where Account_No = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Debtor_Accounts set Account_No = '" & txtSupp.Text & "' where Account_No = '" & txtProd.Text & "'"
        DoEvents
        ActiveReadServer " Update Debtor_Discounts set Debtor_No = '" & txtSupp.Text & "' where Debtor_No = '" & txtProd.Text & "'"
        DoEvents
        frmGuests.txtSuppNo.Text = txtSupp.Text
        frmGuests.grdSupp.TextMatrix(frmGuests.grdSupp.Row, 0) = txtSupp.Text
        Unload Me
    End If
    If KeyCode = 40 Or KeyCode = 38 Then
        KeyCode = 0
        txtProd.SetFocus
    End If
End Sub
Private Sub txtSupp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 32
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub



