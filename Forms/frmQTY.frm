VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmQTY 
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQTY.frx":0000
   ScaleHeight     =   1515
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   450
      TabIndex        =   0
      Top             =   450
      Width           =   3495
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   840
      Index           =   1
      Left            =   4140
      TabIndex        =   1
      Top             =   300
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1482
      Appearance      =   3
      BackColor       =   2163158
      Caption         =   "X"
      CaptionOffsetY  =   2
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmQTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKey_Click(Index As Integer)
    frmStockTake.Tag = ""
    Unload Me
End Sub
Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            frmStockTake.Tag = txtQty.Text
            Unload Me
        Case 27
            frmStockTake.Tag = ""
            Unload Me
    End Select
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
            txtQty.Text = 0
        Case 46, 48 To 57
        Case esle
            KeyAscii = 0
    End Select
End Sub
