VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form cmdInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   Icon            =   "cmdInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   2340
      ScaleHeight     =   2925
      ScaleWidth      =   5355
      TabIndex        =   19
      Top             =   180
      Visible         =   0   'False
      Width           =   5355
      Begin btButtonEx.ButtonEx cmdOk 
         Height          =   315
         Left            =   4110
         TabIndex        =   20
         Top             =   2460
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Appearance      =   3
         Caption         =   "Ok"
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
      Begin MSComCtl2.MonthView mthView 
         Height          =   2310
         Left            =   90
         TabIndex        =   21
         Top             =   90
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16239822
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxSelCount     =   365
         MonthColumns    =   2
         MonthBackColor  =   16777215
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   19660802
         TitleBackColor  =   16761281
         TrailingForeColor=   -2147483639
         CurrentDate     =   38701
      End
      Begin MSForms.Image Image12 
         Height          =   2925
         Left            =   0
         Top             =   0
         Width           =   5355
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "9446;5159"
      End
      Begin MSForms.Image Image11 
         Height          =   2805
         Left            =   60
         Top             =   60
         Width           =   5235
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "9234;4948"
      End
   End
   Begin MSForms.Frame Frame1 
      Height          =   1875
      Left            =   3660
      OleObjectBlob   =   "cmdInfo.frx":000C
      TabIndex        =   16
      Top             =   570
      Width           =   4185
   End
   Begin btButtonEx.ButtonEx cmdSupplier 
      Height          =   315
      Left            =   6750
      TabIndex        =   8
      ToolTipText     =   " Click to Search.... "
      Top             =   2730
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   315
      Left            =   5520
      TabIndex        =   9
      ToolTipText     =   " Click to Search.... "
      Top             =   2730
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   375
      Left            =   7410
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   180
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   8421504
      Caption         =   "¦"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin MSForms.Label lblDate 
      Height          =   315
      Left            =   3720
      TabIndex        =   18
      Top             =   270
      Width           =   3555
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "1 Feb 2006 to 13 Feb 2006"
      Size            =   "6271;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblAv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   2190
      Width           =   1530
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rooms Available:"
      Height          =   195
      Left            =   135
      TabIndex        =   14
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label lblSold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   1875
      Width           =   1530
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rooms Sold:"
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   1890
      Width           =   1695
   End
   Begin VB.Label lblOut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   1545
      Width           =   1530
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Checked Out:"
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Checked In:"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   1230
      Width           =   1695
   End
   Begin VB.Label lblCon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   870
      Width           =   1530
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmed:"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label lblRooms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   210
      Width           =   1530
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Room Quantity:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Provitional Bookings:"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   570
      Width           =   1695
   End
   Begin MSForms.Image Image7 
      Height          =   285
      Left            =   1890
      Top             =   180
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   13827793
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image8 
      Height          =   285
      Left            =   1890
      Top             =   1830
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   13827793
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Left            =   1890
      Top             =   1500
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image3 
      Height          =   285
      Left            =   1890
      Top             =   1170
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image2 
      Height          =   285
      Left            =   1890
      Top             =   840
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image9 
      Height          =   285
      Left            =   1890
      Top             =   2160
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   13827793
      Size            =   "2999;503"
   End
   Begin VB.Label lblProv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   555
      Width           =   1530
   End
   Begin MSForms.Image Image4 
      Height          =   285
      Left            =   1890
      Top             =   510
      Width           =   1700
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "2999;503"
   End
   Begin MSForms.Image Image10 
      Height          =   375
      Left            =   3660
      Top             =   180
      Width           =   3705
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6535;661"
   End
   Begin MSForms.Image Image1 
      Height          =   2445
      Index           =   0
      Left            =   60
      Top             =   90
      Width           =   7875
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "13891;4313"
   End
   Begin MSForms.Image Image6 
      Height          =   2625
      Left            =   30
      Top             =   0
      Width           =   7980
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "14076;4630"
   End
End
Attribute VB_Name = "cmdInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
    Unload Me
End Sub

Private Sub cmdSupplier_Click()
    Unload Me
End Sub

Private Sub lblAcc_Click()

End Sub
