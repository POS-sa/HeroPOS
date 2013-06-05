VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmaddservers 
   Caption         =   "Add Remote Servers"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtport 
      Height          =   315
      Left            =   4260
      TabIndex        =   12
      Text            =   "txtport"
      Top             =   1650
      Width           =   1155
   End
   Begin VB.CommandButton cmdtestcon 
      Caption         =   "Test connection to Selected Server"
      Height          =   465
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add to List"
      Height          =   465
      Left            =   5040
      TabIndex        =   8
      Top             =   2460
      Width           =   1905
   End
   Begin VB.ListBox Listservers 
      Height          =   4155
      Left            =   630
      TabIndex        =   3
      Top             =   510
      Width           =   3495
   End
   Begin VB.TextBox txtpwords 
      Height          =   315
      Left            =   6540
      TabIndex        =   2
      Text            =   "txtpwords"
      Top             =   1650
      Width           =   1215
   End
   Begin VB.TextBox txtusername 
      Height          =   315
      Left            =   6540
      TabIndex        =   1
      Text            =   "txtusername"
      Top             =   810
      Width           =   1665
   End
   Begin VB.TextBox txtserver 
      Height          =   315
      Left            =   4290
      TabIndex        =   0
      Text            =   "txtserver"
      Top             =   810
      Width           =   1875
   End
   Begin VB.Label Label5 
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4290
      TabIndex        =   11
      Top             =   1260
      Width           =   975
   End
   Begin MSForms.CommandButton cmddeleteserver 
      Height          =   405
      Left            =   630
      TabIndex        =   9
      Top             =   5400
      Width           =   3495
      Caption         =   "X Delete Selected Server X"
      Size            =   "6165;714"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6540
      TabIndex        =   7
      Top             =   1260
      Width           =   1425
   End
   Begin VB.Label Label3 
      Caption         =   "Server Admin:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6540
      TabIndex        =   6
      Top             =   420
      Width           =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "Server DNS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4290
      TabIndex        =   5
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Available Servers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   660
      TabIndex        =   4
      Top             =   150
      Width           =   3465
   End
End
Attribute VB_Name = "frmaddservers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
' check for null value
' Add to saved servers and update list

txtserver.Text = ""
txtport.Text = ""
txtusername.Text = ""
txtpwords.Text = ""
End Sub

Private Sub cmddeleteserver_Click()
Dim Response As String
Response = MsgBox("Are you sure you want to delete this Server :  " & ServerName & "  ?", vbYesNo)
Select Case Response

Case vbYes
' Delete the specified server
Case vbNo
'Cancel the deletion and do nothing
End Select
End Sub

Private Sub cmdtestcon_Click()
' connectiontest
End Sub

