VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form cmdReceive 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "cmdReceive.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   840
      Index           =   1
      Left            =   4140
      TabIndex        =   0
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
Attribute VB_Name = "cmdReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub
