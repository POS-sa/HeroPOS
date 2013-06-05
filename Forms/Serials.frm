VERSION 5.00
Begin VB.Form frmSerials 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product verification"
   ClientHeight    =   5295
   ClientLeft      =   6255
   ClientTop       =   3930
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "Serials.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   3750
      Top             =   4680
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1470
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   330
      Top             =   4410
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4260
      Top             =   4680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5850
      MaskColor       =   &H80000010&
      TabIndex        =   11
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5850
      TabIndex        =   10
      Top             =   4830
      Width           =   1095
   End
   Begin VB.TextBox txtperson 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3870
      TabIndex        =   8
      Top             =   1680
      Width           =   3045
   End
   Begin VB.TextBox txttel 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3870
      TabIndex        =   7
      Top             =   1200
      Width           =   3045
   End
   Begin VB.TextBox txtowner 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3870
      TabIndex        =   6
      Top             =   660
      Width           =   3045
   End
   Begin VB.TextBox txt5 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5670
      MaxLength       =   5
      TabIndex        =   4
      Top             =   3645
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txt4 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4320
      MaxLength       =   5
      TabIndex        =   3
      Top             =   3645
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txt3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2970
      MaxLength       =   5
      TabIndex        =   2
      Top             =   3645
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1590
      MaxLength       =   5
      TabIndex        =   1
      Top             =   3645
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   210
      MaxLength       =   5
      TabIndex        =   0
      Top             =   3645
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   5235
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Computer name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   17
      Top             =   2250
      Width           =   3585
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   0
      X2              =   7080
      Y1              =   5250
      Y2              =   5250
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   30
      X2              =   7110
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   7110
      X2              =   7110
      Y1              =   0
      Y2              =   5250
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   30
      X2              =   30
      Y1              =   0
      Y2              =   5250
   End
   Begin VB.Label lblsave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please save all your detail, this will generate a code that you will need to activate HeroPOS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   210
      TabIndex        =   15
      Top             =   2730
      Width           =   2655
   End
   Begin VB.Label lblpcname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3900
      TabIndex        =   14
      Top             =   2250
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact person:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   13
      Top             =   1710
      Width           =   3585
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Telephone number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   12
      Top             =   1200
      Width           =   3585
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Owner/Company name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   9
      Top             =   690
      Width           =   3585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2790
      Width           =   6705
   End
End
Attribute VB_Name = "frmSerials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Reg As RegSearch
Attribute Reg.VB_VarHelpID = -1

Private Sub Form_Initialize()
FindComputerName
If regQuery_A_Key(CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD", "Support") = "" Then Savedowner = False: frmSerials.Tag = "notsaved": Exit Sub
Savedowner = True
'Dim oSecurity As New clsSecurity
'sGetKeyName = Trim("Version1.0.0.858")
'bIsOk = oSecurity.VerifyFile(sGetKeyName, "Registration.dat")
'Set oSecurity = Nothing
'If bIsOk = False Then MsgBox "The registration details was tampered with and is invalid HeroPOS will now close!", vbInformation: End: Unload Me

Checkkey
End Sub

Private Sub Form_Load()
Hd = Str(SNumber("c:\"))
lblpcname.Caption = UCase(ComputerNames)
Screen.MousePointer = 1
Label6.FontSize = 12
Label6.FontBold = True
Label6.Caption = "Welcome to HeroPOS registration."
End Sub
Private Sub Command1_Click()
 Dim regstring As String
 Tstring = txt1.Text & " - " & txt2.Text & " - " & txt3.Text & " - " & txt4.Text & " - " & txt5.Text
 
'    sGetKeyName = Trim("Version1.0.0.858")
'    Dim oSecurity As New clsSecurity
'    Call oSecurity.CreateKeyPair(sGetKeyName)
'    Set oSecurity = Nothing
'    SignFile
'    bIsOk = oSecurity.VerifyFile(sGetKeyName, "Registration.dat")
'    Set oSecurity = Nothing
'
 Snum = txt1.Text & " - " & txt2.Text & " - " & txt3.Text & " - " & txt4.Text & " - " & txt5.Text
 
 
 regCreate_Key_Value HKEY_CURRENT_USER, "Software\HeroPOS\OD", "Registration", Snum, True

Checkkey
End Sub
Private Sub Checkkey()
On Error GoTo ahh
Dim serrval As String, ownerval As String, owners As String, tels As String, pers As String, hds As String
ownerval = regQuery_A_Key(CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD", "Support")
spltowner = Split(ownerval, " - ")
owners = spltowner(0)
tels = spltowner(1)
pers = spltowner(2)
hds = spltowner(3)
hds = LTrim(hds)
Sstring = owners & tels & pers & hds
y = genNumber(Sstring)
serrval = regQuery_A_Key(CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD", "Registration")
x = authKey(serrval, (owners & tels & pers & hds))
If x = True Then
Morne = True


Select Case shw
Case "Main"
frmSerials.Left = 100000
frmMain.Show
Case "Splash"
frmSerials.Left = 100000
frmSplash.Show
Case "Startup"
frmSerials.Left = 100000
Startup.Show
Case ""
frmSerials.Left = 100000
frmMain.Show
End Select
frmSerials.Left = 100000
Exit Sub
End If
On Error Resume Next
'Deleteregistryvalues
MsgBox " An error occured please try to restart the application. If this error continues to happen then please contact HeroPOS for help. "
Unload Me
End

ahh:
Morne = False
Exit Sub
End Sub
Private Sub Deleteregistryvalues()
regDelete_Sub_Key CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD\", "Registration"
regDelete_Sub_Key CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD\", "Support"
x = regDelete_A_Key(CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\", "OD")
x = regDelete_A_Key(CMRegistry.HKEY_CURRENT_USER, "Software", "HeroPOS")
MsgBox "Registration failed please retry or contact HeroPOS support.", vbInformation, "HeroPOS information"
End Sub
Private Sub Command2_Click()
Dim ASS1 As String, Snum As String
ASS1 = LTrim(Hd)
Sstring = txtowner.Text & txttel.Text & txtperson.Text & ASS1
Sstring = UCase(Sstring)
Label1.Caption = "Computer ID: " & ASS1
' txtowner.Tag = genNumber(Sstring)

Dim regstring As String
regstring = txtowner.Text & " - " & txttel.Text & " - " & txtperson.Text & " - " & Hd
regCreate_Key_Value HKEY_CURRENT_USER, "Software\HeroPOS\OD", "Support", regstring, True
Label6.FontSize = 10
Label6.ForeColor = vbRed
Label6.Caption = "Please read all this information accurately back to the HeroPOS support team! "




End Sub

'Private Sub SignFile()
'
'
'Dim oSecurity As New clsSecurity
''
'' Digitally sign a file.
''
'On Error GoTo ErrorHandler
'filenum = FreeFile
'Dim fso As New FileSystemObject
'        If fso.FileExists(App.Path & "\Registration.dat") = False Then fso.CreateTextFile (App.Path & "\Registration.dat")
'        Open Trim(App.Path) & "\Registration.dat" For Output As filenum
'        Dim s As String
'        s = regQuery_A_Key(CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD", "Support")
'            Print #filenum, s, " - ", sGetKeyName, " - ", Snum
'        Close #filenum
'    '
'    ' Sign it.
'    '
'
'        Call oSecurity.SignFile(sGetKeyName, "Registration.dat")
'        MsgBox "Authentication successful!", vbInformation, "Signing complete"
'
'
'Set oSecurity = Nothing
'bIsOk = oSecurity.VerifyFile(sGetKeyName, "Registration.dat")
'Set oSecurity = Nothing
'Exit Sub
'
'ErrorHandler:
'    If err.Number = cdlCancel Then
'        Exit Sub
'    End If
'
'    MsgBox err.Source & vbNewLine & vbNewLine & err.Description

'End Sub





Private Sub Command3_Click()
Screen.MousePointer = 1
Deleteregistryvalues
End
Unload Me
End Sub

Private Sub Timer1_Timer()
If Len(txt5.Text) < 5 Then Exit Sub
If Len(txt4.Text) < 5 Then Exit Sub
If Len(txt3.Text) < 5 Then Exit Sub
If Len(txt2.Text) < 5 Then Exit Sub
If Len(txt1.Text) < 5 Then Exit Sub
Command1.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
If Not regDoes_Key_Exist(CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD") Then Savedowner = False: frmSerials.Tag = "notsaved": Exit Sub
If regQuery_A_Key(CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD", "Support") = "" Then Savedowner = False: frmSerials.Tag = "notsaved": Exit Sub
Savedowner = True
Dim ownerval As String
ownerval = regQuery_A_Key(CMRegistry.HKEY_CURRENT_USER, "Software\HeroPOS\OD", "Support")
splt = Split(ownerval, " - ")
txtowner.Text = splt(0)
txttel.Text = splt(1)
txtperson.Text = splt(2)
frmSerials.Tag = "saved"


Timer1.Enabled = True
txt1.Visible = True
txt2.Visible = True
txt3.Visible = True
txt4.Visible = True
txt5.Visible = True
cmdclear.Visible = True
Timer2.Enabled = False
End Sub

Private Sub FindComputerName()
Screen.MousePointer = 11
Set Reg = New RegSearch
With Reg
    .RootKey = REG_HKEY_LOCAL_MACHINE1
    .SubKey = "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName"
    .SearchFlags = KEY_NAME + VALUE_NAME + VALUE_VALUE + WHOLE_STRING
    .SearchString = "ComputerName"
End With
SearchValue = "FindComputerName"
Reg.DoSearch
End Sub

Private Sub Reg_SearchFound(ByVal sRootKey As String, ByVal sKey As String, ByVal sValue As Variant, ByVal lFound As FOUND_WHERE)
On Error Resume Next
ComputerNames = sValue


'If SearchValue = "FindRegisteredOwner" Then
'    RegisteredOwner = sValue
'ElseIf SearchValue = "FindProductKey" Then
'    ProductKey = sValue
'Else
  
'ElseIf SearchValue = "FindProductName" Then
'    ProductName = sValue

End Sub

Private Sub Timer3_Timer()
    If frmSerials.Tag = "notsaved" Then
        txt1.Visible = False
        txt2.Visible = False
        txt3.Visible = False
        txt4.Visible = False
        txt5.Visible = False
        cmdclear.Visible = False
        Savedowner = False
    
    Else
        Savedowner = True
        txtowner.Enabled = False
        txttel.Enabled = False
        txtperson.Enabled = False
        Command2.Visible = False
        lblsave.Visible = False
        Label1.Visible = True
    End If
End Sub


Private Sub txt1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtowner_GotFocus()
Screen.MousePointer = 3
End Sub

Private Sub txtowner_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtperson_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtperson_LostFocus()
Screen.MousePointer = 1
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
