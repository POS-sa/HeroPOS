VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmChooseRes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose a Reservation to Open"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   Icon            =   "frmChooseRes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdRes 
      Height          =   675
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   90
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1191
      Appearance      =   3
      BackColor       =   12632319
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   150
      Width           =   555
   End
   Begin btButtonEx.ButtonEx cmdRes 
      Height          =   675
      Index           =   1
      Left            =   30
      TabIndex        =   2
      Top             =   870
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1191
      Appearance      =   3
      BackColor       =   16761024
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
End
Attribute VB_Name = "frmChooseRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRes_Click(Index As Integer)
    If frmChooseRes.Caption = "Please Select a Drawer for this Shift." Then
        Select Case Index
            Case 0: frmSplash.cmbUsers.Tag = Mid(cmdRes(Index).Caption, 1, 10)
            Case 1: frmSplash.cmbUsers.Tag = Mid(cmdRes(Index).Caption, 1, 10)
        End Select
        Unload Me
    Else
        If frmChooseRes.Caption = "Select a Reservation" Then
            frmChooseRes.Tag = cmdRes(Index).Tag
            Me.Hide
        Else
            Select Case Index
                Case 0: frmCheckin.lblResNo.Tag = cmdRes(Index).Tag
                Case 1: frmCheckin.lblResNo.Tag = cmdRes(Index).Tag
            End Select
            Unload Me
        End If
    End If
End Sub
Private Sub Form_Activate()
    cmdRes(0).Enabled = True
    cmdRes(1).Enabled = True
            frmChooseRes.cmdRes(0).Caption = "Drawer One"
        frmChooseRes.cmdRes(1).Caption = "Drawer Two"
    ActiveReadServer "Select Drawer_No,User_Name from Users where User_Type = 4 and Workstation_No = " & Workstation_No
    While Not rs.EOF
        Select Case Val(rs.Fields("Drawer_No") & "")
        Case 1: cmdRes(0).Caption = cmdRes(0).Caption & " - " & rs.Fields("User_Name")
        Case 2: cmdRes(1).Caption = cmdRes(1).Caption & " - " & rs.Fields("User_Name")
'            Case 1: cmdRes(0).Enabled = False
'            Case 2: cmdRes(1).Enabled = False
        End Select
        rs.MoveNext
    Wend
    rs.Close
'    If frmChooseRes.Caption = "Please Select a Drawer for this Shift." Then
'        frmChooseRes.cmdRes(0).Caption = "Drawer One"
'        frmChooseRes.cmdRes(1).Caption = "Drawer Two"
'    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmChooseRes.Tag = ""
End Sub
