VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSubscription 
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtauth 
      Height          =   330
      Left            =   3660
      TabIndex        =   6
      Top             =   1980
      Width           =   2865
   End
   Begin VB.TextBox txtsubenddate 
      Height          =   315
      Left            =   8100
      TabIndex        =   4
      Top             =   1980
      Width           =   2895
   End
   Begin btButtonEx.ButtonEx cmdsave 
      Height          =   435
      Left            =   600
      TabIndex        =   0
      Top             =   2820
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Save New  Date and Code"
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
   Begin MSComCtl2.DTPicker dtPicker 
      Height          =   330
      Left            =   600
      TabIndex        =   1
      Top             =   1980
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   65798147
      CurrentDate     =   38776
   End
   Begin btButtonEx.ButtonEx cmdclose 
      Height          =   435
      Left            =   630
      TabIndex        =   9
      Top             =   3510
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Close"
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
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3660
      TabIndex        =   10
      Top             =   2550
      Width           =   2925
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Index           =   1
      Left            =   630
      TabIndex        =   3
      Top             =   330
      Width           =   4725
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Subscription Details"
      Size            =   "8334;556"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label4 
      Caption         =   "Authentication code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3690
      TabIndex        =   8
      Top             =   1710
      Width           =   2085
   End
   Begin VB.Label Label3 
      Caption         =   "Set the date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      TabIndex        =   7
      Top             =   1710
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "Subscription end date - Authenticated"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   8130
      TabIndex        =   5
      Top             =   1680
      Width           =   4125
   End
   Begin MSForms.Image frmTop 
      Height          =   525
      Left            =   480
      Top             =   240
      Width           =   5025
      BackColor       =   16777215
      Size            =   "8864;926"
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2760
      X2              =   12030
      Y1              =   1080
      Y2              =   1080
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   450
      TabIndex        =   2
      Top             =   930
      Width           =   2385
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Subscription Settings"
      Size            =   "4207;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line2 
      X1              =   7170
      X2              =   7170
      Y1              =   1770
      Y2              =   9060
   End
End
Attribute VB_Name = "frmSubscription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
chkdateagain = ""
frmSubscription.Visible = False
Unload frmSubscription
frmDetails.Show
End Sub

Private Sub cmdSave_Click()
If Label5.Caption = "Authenticated" Then
ActiveUpdateServer "Delete from Subscription"
Kill "c:\Windows\Optimumw32.dll"
DoEvents
ActiveUpdateServer "Insert into Subscription (Subscriptiondate, Authcode) values ('" & dtPicker.Value & "','" & txtauth.Text & "')"
DoEvents
If rs.State = 1 Then rs.Close
End If
checked = Checksubscriptiondb
If checked = True Then
txtsubenddate.Text = chkdateagain
End If
End Sub

Private Sub dtPicker_Change()
Dim txtauthtexted As String
Dim dtvalue As String
dtvalue = dtPicker.Value
txtauthtexted = txtauth.Text
If txtauth.Text = "" Then txtauthtexted = "a"
x = dateauthenticate(dtvalue, txtauthtexted)
If x = True Then
Label5.ForeColor = vbGreen
Label5.Caption = "Authenticated"
Else
Label5.ForeColor = vbRed
Label5.Caption = "Not Authorized"
End If
End Sub
Private Sub Form_Activate()
dtPicker.Value = Date
'ActiveReadServer " select * from Subscription"
'If rs.RecordCount > 0 Then
'chkdateagain = Format(rs.Fields("Subscriptiondate"), "DD-MMMM-YYYY")
'txtsubenddate.Text = chkdateagain
'authcodecurrent = rs.Fields("Authcode")
'x = dateauthenticate2(chkdateagain, authcodecurrent)
'End If

checked = Checksubscriptiondb
If checked = True Then
txtsubenddate.Text = chkdateagain
End If
Subscriptloaded = True
End Sub

Public Function dateauthenticate(dtvalue As String, txtauthtext As String) As Boolean
Dim i As String
Dim Thepword As String
Thepword = "subscription"
i = Dcode(txtauthtext, Thepword)
If i = dtPicker.Value Then
dateauthenticate = True
Else
dateauthenticate = False
End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Subscriptloaded = False
chkdateagain = ""
End Sub

Private Sub txtauth_Change()
Dim txtauthtexted As String
Dim dtvalue As String
dtvalue = dtPicker.Value
txtauthtexted = txtauth.Text
If txtauth.Text = "" Then txtauthtexted = "a"
x = dateauthenticate(dtvalue, txtauthtexted)
If x = True Then
Label5.ForeColor = vbGreen
Label5.Caption = "Authenticated"
Else
Label5.ForeColor = vbRed
Label5.Caption = "Not Authorized"
End If
End Sub
