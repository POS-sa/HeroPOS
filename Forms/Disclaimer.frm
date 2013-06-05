VERSION 5.00
Begin VB.Form Disclaimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disclaimer"
   ClientHeight    =   11190
   ClientLeft      =   165
   ClientTop       =   315
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11190
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7515
      Left            =   420
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Disclaimer.frx":0000
      Top             =   1260
      Width           =   13965
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   11340
      TabIndex        =   0
      Top             =   10590
      Width           =   1215
   End
End
Attribute VB_Name = "Disclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Text1.Text = "To the full extent permissible by law, HeroPOS disclaims all responsibility for any damages or losses" & vbCrLf & "(including, without limitation, financial loss, damages for loss in business projects, loss of profits or other consequential losses)" & vbCrLf & "arising in contract, tort or otherwise from the use of or inability to use this Software or any material appearing on www.pos-sa.com, or from any action or decision taken as a result of using HeroPOS or any such material." & vbCrLf & "HeroPOS is not responsible for and has no control over the use of the Software. " & vbCrLf & _
"HeroPOS disclaims all responsibility and liability(including negligence) in relation to information on HeroPOS." & vbCrLf & "By using this software you agree that the Software was bought as is without any contractual binding to change the Software in any way." & vbCrLf & _
"This Software is copyrighted and any illegal copying, reverse engineering or wrapping," & vbCrLf & "will be prosecuted to the fullest extent permissable by law."

End Sub
