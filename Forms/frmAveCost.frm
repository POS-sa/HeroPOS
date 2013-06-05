VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAveCost 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3255
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   2670
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   767
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSForms.Image Image1 
      Height          =   1725
      Index           =   0
      Left            =   -30
      Top             =   -30
      Width           =   7005
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   3
      Size            =   "12356;3043"
      Picture         =   "frmAveCost.frx":0000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recalculate Average Cost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   390
      TabIndex        =   2
      Top             =   1740
      Width           =   6135
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Top             =   2400
      Width           =   5265
   End
   Begin MSForms.Image Image1 
      Height          =   1125
      Index           =   1
      Left            =   -30
      Top             =   2160
      Width           =   7005
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;1984"
   End
   Begin MSForms.Image Image1 
      Height          =   555
      Index           =   2
      Left            =   -30
      Top             =   1650
      Width           =   7005
      BackColor       =   15523287
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;979"
   End
End
Attribute VB_Name = "frmAveCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    DoEvents
    Recalc
End Sub
Private Sub Recalc()
    On Error Resume Next
    ProgressBar1.Value = 0
    ActiveReadServer "Select Product_Code,Description from Products order by Product_Code"
    ProgressBar1.Max = rs.RecordCount
    While Not rs.EOF
        ProgressBar1.Value = ProgressBar1.Value + 1
        DoEvents
        lblCap.Caption = "Recalculating " & rs.Fields("Product_Code") & " - " & rs.Fields("Description")
        ActiveUpdateServer "Update Products set Ave_Cost = Landed_Cost "
        'Where Product_Code = " & rs.Fields("Product_Code")
        rs.MoveNext
    Wend
    rs.Close
    ProgressBar1.Value = ProgressBar1.Max
    MsgBox "Recalulation Completed Succesfully"
    Unload Me
    On Error Resume Next
End Sub

