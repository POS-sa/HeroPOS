VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDepChange 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Department Number Change..."
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   630
      Left            =   3975
      TabIndex        =   2
      Top             =   90
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
   Begin VB.TextBox txtProd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   90
      Width           =   2715
   End
   Begin VB.TextBox txtSupp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1215
      TabIndex        =   1
      Top             =   420
      Width           =   2715
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Number:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "New Number:"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   450
      Width           =   1065
   End
End
Attribute VB_Name = "frmDepChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
    If txtProd.Text = "" Or txtSupp.Text = "" Then
        MsgBox "Blank Department Numbers are not Allowed", vbCritical, "HeroPOS"
        Exit Sub
    End If
    Screen.MousePointer = 11
    ActiveUpdateServer "Update Departments set Department_No = '" & txtSupp.Text & "' where Department_No = '" & txtProd.Text & "'"
    DoEvents
    ActiveUpdateServer "Update Products set Department_No = '" & txtSupp.Text & "' where Department_No = '" & txtProd.Text & "'"
    DoEvents
    ActiveUpdateServer "Update Department_Links set Dept_No = '" & txtSupp.Text & "' where Dept_No = '" & txtProd.Text & "'"
    DoEvents
    ActiveUpdateServer "Update Sales_Journal set Department_No = '" & txtSupp.Text & "' where Department_No = '" & txtProd.Text & "'"
    Screen.MousePointer = 0
    Unload Me
End Sub
Private Sub Form_Load()
    With frmDepartments
        If .picMaj.BackColor = &HFFC0C0 Then
            If .grdMajor.Rows > 1 Then
                txtProd.Text = .grdMajor.TextMatrix(.grdMajor.Row, 0)
            End If
        End If
        If .picMin.BackColor = &HFFC0C0 Then
            If .grdSub.Rows > 1 Then
                txtProd.Text = .grdSub.TextMatrix(.grdSub.Row, 0)
            End If
        End If
    End With
End Sub
