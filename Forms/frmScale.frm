VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmScale 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmScale.frx":0000
   ScaleHeight     =   9285
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   1740
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   420
      Top             =   3270
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4230
      ScaleHeight     =   315
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   1320
      Width           =   615
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   840
      Index           =   1
      Left            =   4170
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   735
      _ExtentX        =   1296
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
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1110
      Index           =   0
      Left            =   330
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2850
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1958
      Appearance      =   3
      BackColor       =   16761024
      Caption         =   "Accept"
      CaptionOffsetY  =   2
      ForeColor       =   9197376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   0
      Left            =   470
      TabIndex        =   7
      Top             =   4550
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   4550
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   2
      Left            =   3370
      TabIndex        =   9
      Top             =   4550
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   3
      Left            =   470
      TabIndex        =   10
      Top             =   5680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   4
      Left            =   1920
      TabIndex        =   11
      Top             =   5680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   5
      Left            =   3370
      TabIndex        =   12
      Top             =   5680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   6
      Left            =   470
      TabIndex        =   13
      Top             =   6810
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   7
      Left            =   1920
      TabIndex        =   14
      Top             =   6810
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   8
      Left            =   3370
      TabIndex        =   15
      Top             =   6810
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   9
      Left            =   470
      TabIndex        =   16
      Top             =   7940
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "CL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   10
      Left            =   1920
      TabIndex        =   17
      Top             =   7940
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   11
      Left            =   3370
      TabIndex        =   18
      Top             =   7940
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.TextBox lblWeightType 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   1335
      Left            =   368
      TabIndex        =   6
      Top             =   1350
      Visible         =   0   'False
      Width           =   3585
      VariousPropertyBits=   746604561
      Size            =   "6324;2355"
      Value           =   "0.00"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073750016
      FontHeight      =   1125
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   840
      Left            =   3950
      TabIndex        =   4
      Top             =   1710
      Width           =   960
   End
   Begin VB.Label lblWeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   56.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   330
      TabIndex        =   3
      Top             =   1380
      Width           =   3585
   End
   Begin VB.Label lblPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   480
      TabIndex        =   2
      Top             =   250
      Width           =   2700
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      BorderWidth     =   10
      FillColor       =   &H00CDD1D0&
      FillStyle       =   0  'Solid
      Height          =   4900
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   4335
      Width           =   5085
   End
End
Attribute VB_Name = "frmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInput_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 11
            If Me.lblWeightType <> "0.00" Then
                Me.lblWeightType.Text = Me.lblWeightType.Text & Me.cmdInput(Index).Caption
            Else
                Me.lblWeightType.Text = Me.cmdInput(Index).Caption
            End If
        Case 9
            Me.lblWeightType = "0.00"
            Call lblWeightType_GotFocus
    End Select
End Sub

Private Sub cmdKey_Click(Index As Integer)
    Select Case Index
        Case 0
            TillData.Weight = lblWeight.Caption
            Unload Me
        Case 1
            TillData.Weight = 0
            Unload Me
    End Select
End Sub

Private Sub Form_Initialize()
    Inside = False
End Sub

Private Sub Form_Load()
    Inside = False
    If Me.lblWeight.Visible = False Then
        Screen.MousePointer = vbDefault
        Me.Height = 4260
    End If
    Me.Shape1.BackColor = RGB(62, 93, 122)
    Me.Shape1.BorderColor = RGB(128, 150, 173)
    Me.BackColor = RGB(62, 93, 122)
End Sub

Private Sub lblWeightType_Change()
    lblWeight.Caption = Me.lblWeightType.Text
    lblPrice = Format(TillData.Price * Val(lblWeight.Caption), "0.00")
End Sub

Private Sub lblWeightType_GotFocus()
    If Me.lblWeight.Visible = False Then
        Screen.MousePointer = vbDefault
        With Me.lblWeightType
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub

Private Sub lblWeightType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57
        Case 46
            If Len(Me.lblWeightType.Text) <> 0 Then
                If InStr(1, Me.lblWeightType.Text, ".", vbTextCompare) <> 0 Then
                    KeyAscii = 0
                End If
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Timer1_Timer()
    If Me.lblWeight.Visible = False Then
        Me.Timer1.Enabled = False
        Exit Sub
    End If
  
    On Error GoTo far
    If Inside = True Then Exit Sub
    Inside = True
    stime = Timer
    
    MSComm1.CommPort = Right(Devices.ScalePort, 1)
    MSComm1.InBufferSize = 20
    MSComm1.Settings = Devices.ScaleSet
    MSComm1.InputLen = 0
    MSComm1.Inputmode = comInputModeText
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
    MSComm1.Output = Chr(5)
   
    While Timer - stime < 0.4: Wend
    buffer$ = ""
    i = 0
    Do
        i = i + 1
        buffer$ = buffer$ & MSComm1.Input
    Loop Until InStr(buffer$, Chr(30)) Or i = 20
    If buffer$ <> "" Then
        lblWeight.Caption = Format(Val(Mid(buffer$, 1, Len(buffer$) - 1)) / 1000, "0.000")
        lblPrice = Format(TillData.Price * Val(lblWeight.Caption), "0.00")
    Else
        If i = 20 Then lblWeight.Caption = "0.00"
    End If
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    While Timer - stime < 0.01: Wend
    Inside = False
    Exit Sub
far:
    On Error GoTo 0
End Sub
