VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmCheckin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guest Check In"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   FillColor       =   &H000000C0&
   Icon            =   "frmCheckin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCredit 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8310
      TabIndex        =   81
      Top             =   4950
      Width           =   2565
   End
   Begin VB.TextBox txtFree 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10140
      TabIndex        =   21
      Text            =   "0"
      Top             =   1290
      Width           =   735
   End
   Begin VB.PictureBox picRates 
      Height          =   315
      Left            =   8160
      ScaleHeight     =   255
      ScaleWidth      =   2685
      TabIndex        =   74
      Top             =   3420
      Visible         =   0   'False
      Width           =   2745
      Begin VB.Label lblRatestring 
         Height          =   165
         Left            =   150
         TabIndex        =   75
         Top             =   30
         Width           =   2325
      End
   End
   Begin VB.TextBox txtContactNo 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8310
      TabIndex        =   28
      Top             =   3120
      Width           =   2565
   End
   Begin VB.TextBox txtContact 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8310
      TabIndex        =   27
      Top             =   2760
      Width           =   2565
   End
   Begin VB.TextBox txt5 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10140
      TabIndex        =   23
      Text            =   "0"
      Top             =   1680
      Width           =   765
   End
   Begin VB.TextBox txtAdults 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8310
      TabIndex        =   22
      Text            =   "0"
      Top             =   1680
      Width           =   765
   End
   Begin VB.TextBox txt0to5 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8310
      TabIndex        =   24
      Text            =   "0"
      Top             =   2040
      Width           =   765
   End
   Begin VB.TextBox txt12to16 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10140
      TabIndex        =   25
      Text            =   "0"
      Top             =   2010
      Width           =   765
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   8310
      TabIndex        =   20
      Top             =   1290
      Width           =   645
   End
   Begin VB.TextBox txtNights 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   1065
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1680
      TabIndex        =   8
      Top             =   2400
      Width           =   3165
   End
   Begin VB.TextBox txtFax 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1710
      TabIndex        =   15
      Top             =   4920
      Width           =   3165
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   4560
      Width           =   3165
   End
   Begin VB.TextBox txtTel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   4200
      Width           =   3165
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   3840
      Width           =   4275
   End
   Begin VB.TextBox txtVehReg 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   3480
      Width           =   2625
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   3120
      Width           =   2595
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   2760
      Width           =   4275
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1770
      TabIndex        =   1
      Top             =   540
      Width           =   4185
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   315
      Left            =   9240
      TabIndex        =   35
      Top             =   7110
      Width           =   1665
      _ExtentX        =   2937
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
      ShowFocus       =   0
   End
   Begin MSComCtl2.DTPicker DTArrTime 
      Height          =   345
      Left            =   3660
      TabIndex        =   3
      Top             =   840
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Format          =   68157442
      CurrentDate     =   38862.5833333333
   End
   Begin MSComCtl2.DTPicker DTDepTime 
      Height          =   345
      Left            =   3660
      TabIndex        =   5
      Top             =   1230
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Format          =   68157442
      CurrentDate     =   38862.4583333333
   End
   Begin MSComCtl2.DTPicker DTDeparture 
      Height          =   345
      Left            =   1620
      TabIndex        =   4
      Top             =   1230
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   68157443
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTArrival 
      Height          =   345
      Left            =   1620
      TabIndex        =   2
      Top             =   840
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   68157443
      CurrentDate     =   38862
   End
   Begin btButtonEx.ButtonEx cmdAccept 
      Height          =   315
      Left            =   7560
      TabIndex        =   34
      Top             =   7110
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Accept"
      Enabled         =   0   'False
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
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   1155
      Left            =   1740
      TabIndex        =   16
      Top             =   5280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2037
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmCheckin.frx":000C
   End
   Begin btButtonEx.ButtonEx cmdCondition 
      Height          =   315
      Left            =   2880
      TabIndex        =   65
      Top             =   105
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Conditions of Stay..."
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":008E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":05B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":0AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":0FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":1516
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":1A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":1F5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":247C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":299E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":2EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":33E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":3904
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":3E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":4348
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":486A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":4D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":52AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":57D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":5CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":6214
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":6736
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":6C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":717A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":769C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":7BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":80E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":8602
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":8B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckin.frx":9046
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtRemarks 
      Height          =   1155
      Left            =   8310
      TabIndex        =   33
      ToolTipText     =   " Check In Task "
      Top             =   5310
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2037
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmCheckin.frx":9568
   End
   Begin BTNENHLib4.BtnEnh fmType 
      Height          =   1005
      Left            =   5070
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2955
      _Version        =   524298
      _ExtentX        =   5212
      _ExtentY        =   1773
      _StockProps     =   66
      Caption         =   "Provisional Booking"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      CornerFactor    =   10
      Surface         =   3
      PictureTranspColor=   192
      BackColorContainer=   14215660
      ShadowColor     =   16777215
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      ForeColorDisabled=   12640511
      UserData        =   0.1
      textCaption     =   "frmCheckin.frx":95EA
      textLT          =   "frmCheckin.frx":9670
      textCT          =   "frmCheckin.frx":9688
      textRT          =   "frmCheckin.frx":96A0
      textLM          =   "frmCheckin.frx":96B8
      textRM          =   "frmCheckin.frx":96D0
      textLB          =   "frmCheckin.frx":96E8
      textCB          =   "frmCheckin.frx":9700
      textRB          =   "frmCheckin.frx":9718
      colorBack       =   "frmCheckin.frx":9730
      colorIntern     =   "frmCheckin.frx":975A
      colorMO         =   "frmCheckin.frx":9784
      colorFocus      =   "frmCheckin.frx":97AE
      colorDisabled   =   "frmCheckin.frx":97D8
      colorPressed    =   "frmCheckin.frx":9802
      HollowFrame     =   -1  'True
      LightDirection  =   8
   End
   Begin btButtonEx.ButtonEx cmdDeposit 
      Height          =   315
      Left            =   5820
      TabIndex        =   70
      Top             =   7110
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Receive Deposit"
      Enabled         =   0   'False
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
   Begin btButtonEx.ButtonEx cmdRate 
      Height          =   300
      Left            =   6750
      TabIndex        =   72
      Top             =   3420
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Applied Rates..."
      Enabled         =   0   'False
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
   Begin btButtonEx.ButtonEx cmdGroup 
      Height          =   315
      Left            =   4590
      TabIndex        =   80
      Top             =   105
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Group..."
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
   Begin MSForms.CheckBox CheckBox1 
      Height          =   285
      Left            =   8040
      TabIndex        =   83
      Top             =   6720
      Width           =   2835
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "5001;503"
      Value           =   "0"
      Caption         =   "Save Client Details to Adress Book"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card Details:"
      Height          =   255
      Left            =   6690
      TabIndex        =   82
      Top             =   4950
      Width           =   1395
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   19
      Left            =   8160
      Top             =   4890
      Width           =   2745
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "4842;556"
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Com Nights:"
      Height          =   255
      Left            =   8910
      TabIndex        =   79
      Top             =   1290
      Width           =   1035
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   18
      Left            =   9990
      Top             =   1200
      Width           =   915
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1614;609"
   End
   Begin VB.Label lblTotRoom 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0.00"
      Height          =   225
      Left            =   9300
      TabIndex        =   78
      Top             =   3855
      Width           =   1545
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Room Balance:"
      Height          =   165
      Left            =   2670
      TabIndex        =   77
      Top             =   7170
      Width           =   1095
   End
   Begin VB.Label lblBal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3930
      TabIndex        =   76
      Top             =   7170
      Width           =   1725
   End
   Begin VB.Label lblTotRate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0.00"
      Height          =   195
      Left            =   8340
      TabIndex        =   30
      Top             =   3855
      Width           =   675
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rate:"
      Height          =   255
      Left            =   6690
      TabIndex        =   73
      Top             =   3840
      Width           =   1395
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   16
      Left            =   8160
      Top             =   3780
      Width           =   1005
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1773;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   15
      Left            =   8160
      Top             =   3060
      Width           =   2745
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "4842;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   14
      Left            =   8160
      Top             =   2700
      Width           =   2745
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "4842;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   13
      Left            =   9990
      Top             =   1590
      Width           =   915
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1614;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   12
      Left            =   8160
      Top             =   1980
      Width           =   915
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1614;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   11
      Left            =   8160
      Top             =   1620
      Width           =   915
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1614;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   10
      Left            =   9990
      Top             =   1950
      Width           =   915
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1614;556"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   9
      Left            =   8160
      Top             =   1230
      Width           =   795
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1402;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   8
      Left            =   1620
      Top             =   4860
      Width           =   3315
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "5847;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   7
      Left            =   1620
      Top             =   4500
      Width           =   3345
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "5900;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   6
      Left            =   1620
      Top             =   4140
      Width           =   3345
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "5900;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   5
      Left            =   1620
      Top             =   3780
      Width           =   4455
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "7858;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   4
      Left            =   1620
      Top             =   3420
      Width           =   2775
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "4895;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   3
      Left            =   1620
      Top             =   3060
      Width           =   2775
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "4895;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   2
      Left            =   1620
      Top             =   2700
      Width           =   4455
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "7858;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Index           =   1
      Left            =   1620
      Top             =   2340
      Width           =   3345
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "5900;609"
   End
   Begin MSForms.Image Image6 
      Height          =   315
      Left            =   1620
      Top             =   1620
      Width           =   1215
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "2143;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   0
      Left            =   1620
      Top             =   480
      Width           =   4455
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "7858;556"
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Fax No:"
      Height          =   255
      Left            =   150
      TabIndex        =   71
      Top             =   4920
      Width           =   1395
   End
   Begin MSForms.Image Image4 
      Height          =   1305
      Left            =   8160
      Top             =   5250
      Width           =   2745
      BorderColor     =   0
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "4842;2302"
   End
   Begin VB.Label lblUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   660
      TabIndex        =   68
      Top             =   7170
      Width           =   1725
   End
   Begin MSForms.Image Image2 
      Height          =   315
      Left            =   570
      Top             =   7110
      Width           =   1875
      BackColor       =   16777215
      Size            =   "3307;556"
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      Height          =   165
      Left            =   -330
      TabIndex        =   67
      Top             =   7170
      Width           =   855
   End
   Begin VB.Label lblResNo 
      Height          =   315
      Left            =   5430
      TabIndex        =   66
      Top             =   870
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Image Image3 
      Height          =   1305
      Left            =   1620
      Top             =   5220
      Width           =   3315
      BorderColor     =   0
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "5847;2302"
      VariousPropertyBits=   25
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Address:"
      Height          =   195
      Left            =   330
      TabIndex        =   64
      Top             =   5220
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Required fields = *"
      Height          =   255
      Left            =   120
      TabIndex        =   63
      Top             =   6330
      Width           =   1545
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      Height          =   315
      Left            =   6150
      TabIndex        =   62
      Top             =   5250
      Width           =   1965
   End
   Begin MSForms.ComboBox cmbPay 
      Height          =   315
      Left            =   8160
      TabIndex        =   32
      Tag             =   "Up"
      Top             =   4530
      Width           =   2745
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Responsibility:"
      Height          =   255
      Left            =   6150
      TabIndex        =   61
      Top             =   4590
      Width           =   1935
   End
   Begin MSForms.ComboBox cmbBusiness 
      Height          =   315
      Left            =   8160
      TabIndex        =   31
      Tag             =   "Up"
      Top             =   4170
      Width           =   2745
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business Source:"
      Height          =   255
      Left            =   6690
      TabIndex        =   60
      Top             =   4230
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbRate 
      Height          =   315
      Left            =   8160
      TabIndex        =   29
      Tag             =   "Up"
      Top             =   3420
      Width           =   2745
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
      Height          =   255
      Left            =   6690
      TabIndex        =   59
      Top             =   3120
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbBooked 
      Height          =   315
      Left            =   8160
      TabIndex        =   26
      Tag             =   "Up"
      Top             =   2340
      Width           =   2745
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   7
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Booked by:"
      Height          =   255
      Left            =   6690
      TabIndex        =   58
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "12-16 yrs:"
      Height          =   255
      Left            =   9210
      TabIndex        =   57
      Top             =   1980
      Width           =   735
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0-5 yrs:"
      Height          =   255
      Left            =   7380
      TabIndex        =   56
      Top             =   2070
      Width           =   735
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person:"
      Height          =   255
      Left            =   6690
      TabIndex        =   55
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5-12 yrs:"
      Height          =   225
      Left            =   8550
      TabIndex        =   54
      Top             =   1710
      Width           =   1395
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Adults:"
      Height          =   255
      Left            =   6690
      TabIndex        =   53
      Top             =   1650
      Width           =   1395
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      Height          =   255
      Left            =   6690
      TabIndex        =   52
      Top             =   1290
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbCity 
      Height          =   345
      Left            =   8160
      TabIndex        =   19
      Tag             =   "Up"
      Top             =   840
      Width           =   2745
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4842;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "Cape Town"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City/Town:"
      Height          =   255
      Left            =   6690
      TabIndex        =   51
      Top             =   900
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbProvince 
      Height          =   315
      Left            =   8160
      TabIndex        =   18
      Tag             =   "Up"
      Top             =   480
      Width           =   2745
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4842;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "Western Cape"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Province:"
      Height          =   255
      Left            =   6690
      TabIndex        =   50
      Top             =   540
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   345
      Left            =   8160
      TabIndex        =   17
      Tag             =   "Up"
      Top             =   90
      Width           =   2745
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4842;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Country of Origin:"
      Height          =   255
      Left            =   6690
      TabIndex        =   49
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No:"
      Height          =   255
      Left            =   150
      TabIndex        =   48
      Top             =   4560
      Width           =   1395
   End
   Begin MSForms.ComboBox cmbRoomNo 
      Height          =   345
      Left            =   1620
      TabIndex        =   0
      Tag             =   "Up"
      Top             =   90
      Width           =   1215
      VariousPropertyBits=   746604571
      MaxLength       =   4
      DisplayStyle    =   7
      Size            =   "2143;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbTitle 
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Tag             =   "Up"
      Top             =   1980
      Width           =   1245
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2196;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Tel No:"
      Height          =   255
      Left            =   150
      TabIndex        =   47
      Top             =   4200
      Width           =   1395
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   150
      TabIndex        =   46
      Top             =   3840
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Veh Reg No:"
      Height          =   255
      Left            =   150
      TabIndex        =   45
      Top             =   3480
      Width           =   1395
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*ID No:"
      Height          =   255
      Left            =   150
      TabIndex        =   44
      Top             =   3120
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Last Name:"
      Height          =   255
      Left            =   150
      TabIndex        =   43
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*First Name:"
      Height          =   255
      Left            =   150
      TabIndex        =   42
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nights:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   41
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Title:"
      Height          =   255
      Left            =   150
      TabIndex        =   40
      Top             =   2070
      Width           =   1395
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   39
      Top             =   180
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Room Number:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Index           =   12
      Left            =   180
      TabIndex        =   38
      Top             =   540
      Width           =   1395
      VariousPropertyBits=   8388627
      Caption         =   " Room Description:"
      Size            =   "2461;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   37
      Top             =   900
      Width           =   1395
      Caption         =   "*Arrival Date:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   36
      Top             =   1290
      Width           =   1395
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "*Departure Date:"
      Size            =   "2461;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image7 
      Height          =   315
      Left            =   3840
      Top             =   7110
      Width           =   1875
      BackColor       =   16777215
      Size            =   "3307;556"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   17
      Left            =   9210
      Top             =   3780
      Width           =   1695
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "2990;556"
   End
   Begin MSForms.Image Image1 
      Height          =   6705
      Left            =   30
      Top             =   -30
      Width           =   11025
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "19447;11827"
   End
End
Attribute VB_Name = "frmCheckin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calc_Rate()
    picRates.Visible = False
    cmbRate.Enabled = True
    cmdRate.Enabled = False
    txtNights.Text = Val(DTDeparture.Value - DTArrival.Value)
    lblRatestring = ""
    If Val(txtAdults.Text) <> 0 Then
        lblTotRate.Caption = Format(Val(Mid(cmbRate.Text, InStr(cmbRate.Text, ">") + 1)), "0.00")
        lblRatestring.Caption = Trim(Mid(cmbRate.Text, 1, InStr(cmbRate.Text, "-") - 1))
    Else
        lblTotRate.Caption = "0.00"
        lblRatestring.Caption = ""
    End If
    If Val(txt5.Text) <> 0 Then
        ActiveReadServer2 "Select * from Room_Rates where Condition = 'Children 5 to 12'"
        If rs2.RecordCount > 0 Then
            lblRatestring.Caption = lblRatestring.Caption & "-" & rs2.Fields("Rate_Type")
            lblTotRate.Caption = Format(Val(lblTotRate.Caption) + (Val(txt5.Text) * rs2.Fields("Room_Rate")), "0.00")
        End If
        rs2.Close
    End If
    If Val(txt0to5.Text) <> 0 Then
        ActiveReadServer2 "Select * from Room_Rates where Condition = 'Children under 5'"
        If rs2.RecordCount > 0 Then
            lblRatestring.Caption = lblRatestring.Caption & "-" & rs2.Fields("Rate_Type")
            lblTotRate.Caption = Format(Val(lblTotRate.Caption) + (Val(txt0to5.Text) * rs2.Fields("Room_Rate")), "0.00")
        End If
        rs2.Close
    End If
    If Val(txt12to16.Text) <> 0 Then
        ActiveReadServer2 "Select * from Room_Rates where Condition = 'Children 12 to 15'"
        If rs2.RecordCount > 0 Then
            lblRatestring.Caption = lblRatestring.Caption & "-" & rs2.Fields("Rate_Type")
            lblTotRate.Caption = Format(Val(lblTotRate.Caption) + (Val(txt12to16.Text) * rs2.Fields("Room_Rate")), "0.00")
        End If
        rs2.Close
    End If
    If InStr(lblRatestring.Caption, "-") <> 0 Then
        cmdRate.Enabled = True
        cmbRate.Enabled = False
        picRates.Visible = True
    End If
    lblTotRoom.Caption = (Val(frmCheckin.txtNights.Text) - Val(frmCheckin.txtFree.Text)) & " x Nights = " & Format(Val(frmCheckin.lblTotRate.Caption) * (Val(frmCheckin.txtNights.Text) - Val(frmCheckin.txtFree.Text)), "0.00")
    Debug.Print (Val(frmCheckin.txtNights.Text) - Val(frmCheckin.txtFree.Text)) & " x Nights = " & Format(Val(frmCheckin.lblTotRate.Caption) * (Val(frmCheckin.txtNights.Text) - Val(frmCheckin.txtFree.Text)), "0.00")
    
    
End Sub
Private Function Validate()
    Validate = True
    If cmbTitle.Text = "" Then Validate = False
    If txtFName.Text = "" Then Validate = False
    If txtLName.Text = "" Then Validate = False
    If txtTel.Text = "" Then Validate = False
    If txtFax.Text = "" Then Validate = False
    If txtAddress.Text = "" Then Validate = False
    If Val(txtAdults.Text) + Val(txt5.Text) + Val(txt0to5.Text) + Val(txt12to16.Text) = 0 Then Validate = False
    If fmType.Caption = "Guest Checked Out" Then Validate = False
    If fmType.Caption = "Guest Checked In" Then Validate = True
End Function
Private Sub LoadResDetail(Res_No)
fmType.Tag = "1"
On Error Resume Next
    cmbRate.Enabled = False
    cmbRoomNo.Enabled = False
    txtDescription.Enabled = False
    DTArrival.Enabled = False
    DTDeparture.Enabled = False
    cmbTitle.Enabled = False
    txtFName.Enabled = False
    txtLName.Enabled = False
    txtID.Enabled = False
    txtVehReg.Enabled = False
    txtEmail.Enabled = False
    txtTel.Enabled = False
    txtMobile.Enabled = False
    txtFax.Enabled = False
    txtAddress.Enabled = False
    txtCode.Enabled = False
    txtAdults.Enabled = False
    txtFree.Enabled = False
    cmbCountry.Enabled = False
    cmbProvince.Enabled = False
    cmbCity.Enabled = False
    txt5.Enabled = False
    txt0to5.Enabled = False
    txt12to16.Enabled = False
    txtContact.Enabled = False
    txtContactNo.Enabled = False
    txtRemarks.Enabled = False
    cmbBusiness.Enabled = False
    cmbPay.Enabled = False
    cmbBooked.Enabled = False
    txtNights.Enabled = False
    DTArrTime.Enabled = False
    DTDepTime.Enabled = False
    txtCredit.Enabled = False
    Image3.BackColor = &H8000000F
    Image4.BackColor = &H8000000F
    Action = "None"
    For i = 0 To frmRes.cmdTop.Count - 1
        If frmRes.cmdTop(i).FontBold = True Then
            Action = frmRes.cmdTop(i).Caption
            Exit For
        End If
    Next i
    ActiveReadServer "Select * from Reservations where Res_no = " & Res_No
    lblResNo.Caption = Res_No
    If rs.RecordCount > 0 Then
        txtAdults.Text = rs.Fields("Adults") & ""
        txt5.Text = Val(rs.Fields("Kid5to12") & "")
        txt0to5.Text = Val(rs.Fields("Kid0to5") & "")
        txt12to16.Text = Val(rs.Fields("Kids12to16") & "")
        For i = 0 To cmbRate.ListCount - 1
            If Val(Mid(cmbRate.List(i), 1, InStr(cmbRate.List(i), "-") - 1)) = rs.Fields("Rate_Type") Then
                cmbRate.ListIndex = rs.Fields("Rate_Type") - 1
                Exit For
            End If
        Next i
        Select Case rs.Fields("Res_Type")
            Case 0
                frmCheckin.Caption = "Guest Check In > Reservation No: " & Res_No
                fmType.ForeColor = &HC0C0&
                fmType.Caption = "Provisional Booking"
                cmbRate.Enabled = True
                txtFree.Enabled = True
                cmdDeposit.Enabled = True
                DTArrTime.Enabled = True
                DTDepTime.Enabled = True
                cmbTitle.Enabled = True
                txtFName.Enabled = True
                txtLName.Enabled = True
                txtID.Enabled = True
                txtVehReg.Enabled = True
                txtEmail.Enabled = True
                txtTel.Enabled = True
                txtMobile.Enabled = True
                txtFax.Enabled = True
                txtAddress.Enabled = True
                txtCode.Enabled = True
                txtAdults.Enabled = True
                cmbCountry.Enabled = True
                cmbProvince.Enabled = True
                cmbCity.Enabled = True
                txt5.Enabled = True
                txt0to5.Enabled = True
                txt12to16.Enabled = True
                txtContact.Enabled = True
                txtContactNo.Enabled = True
                txtRemarks.Enabled = True
                cmbBusiness.Enabled = True
                cmbPay.Enabled = True
                cmbBooked.Enabled = True
                txtCredit.Enabled = True
                Image3.BackColor = &HFFFFFF
                Image4.BackColor = &HFFFFFF
                txtNights.Enabled = False
                If cmdAccept.Caption = "Save" Then
                    DTArrival.Enabled = True
                    DTDeparture.Enabled = True
                    cmbRoomNo.Enabled = True
                End If
                If cmdAccept.Caption = "" Then cmdAccept.Caption = "Check In"
            Case 1
                frmCheckin.Caption = "Guest Check In > Reservation No: " & Res_No
                fmType.ForeColor = &H8000&
                fmType.Caption = "Confirmed Booking"
                cmbRate.Enabled = True
                cmdDeposit.Enabled = True
                DTArrTime.Enabled = True
                txtFree.Enabled = True
                DTDepTime.Enabled = True
                cmbTitle.Enabled = True
                txtFName.Enabled = True
                txtLName.Enabled = True
                txtID.Enabled = True
                txtVehReg.Enabled = True
                txtEmail.Enabled = True
                txtTel.Enabled = True
                txtMobile.Enabled = True
                txtFax.Enabled = True
                txtAddress.Enabled = True
                txtCode.Enabled = True
                txtAdults.Enabled = True
                cmbCountry.Enabled = True
                cmbProvince.Enabled = True
                cmbCity.Enabled = True
                txt5.Enabled = True
                txt0to5.Enabled = True
                txt12to16.Enabled = True
                txtContact.Enabled = True
                txtContactNo.Enabled = True
                txtRemarks.Enabled = True
                cmbBusiness.Enabled = True
                cmbPay.Enabled = True
                txtCredit.Enabled = True
                cmbBooked.Enabled = True
                Image3.BackColor = &HFFFFFF
                Image4.BackColor = &HFFFFFF
                txtNights.Enabled = False
                If cmdAccept.Caption = "" Then cmdAccept.Caption = "Check In"
                If cmdAccept.Caption = "Save" Then
                    DTArrival.Enabled = True
                    DTDeparture.Enabled = True
                    cmbRoomNo.Enabled = True
                    cmbTitle.Enabled = True
                    txtFName.Enabled = True
                    txtLName.Enabled = True
                    txtID.Enabled = True
                    txtVehReg.Enabled = True
                    txtEmail.Enabled = True
                    txtTel.Enabled = True
                    txtMobile.Enabled = True
                    txtFax.Enabled = True
                    txtAddress.Enabled = True
                    txtCode.Enabled = True
                    cmbCountry.Enabled = True
                    cmbProvince.Enabled = True
                    cmbCity.Enabled = True
                    Image3.BackColor = &HFFFFFF
                End If
            Case 2
                cmbRate.Enabled = False
                frmCheckin.Caption = "Guest Check In > Reservation No: " & Res_No
                fmType.ForeColor = &HC00000
                fmType.Caption = "Guest Checked In"
                cmdDeposit.Enabled = False
                If cmdAccept.Caption = "" Then cmdAccept.Caption = "Check Out"
                If cmdAccept.Caption = "Save" Then
                    DTDeparture.Enabled = True
                    cmbTitle.Enabled = True
                    txtFName.Enabled = True
                    txtLName.Enabled = True
                    txtID.Enabled = True
                    txtVehReg.Enabled = True
                    txtEmail.Enabled = True
                    txtTel.Enabled = True
                    txtMobile.Enabled = True
                    txtFax.Enabled = True
                    txtAddress.Enabled = True
                    txtCode.Enabled = True
                    cmbCountry.Enabled = True
                    cmbProvince.Enabled = True
                    txtCredit.Enabled = True
                    cmbCity.Enabled = True
                    Image3.BackColor = &HFFFFFF
                End If
            Case 3
                cmbRate.Enabled = False
                frmCheckin.Caption = "Guest Check In > Reservation No: " & Res_No
                fmType.ForeColor = &HC0&
                fmType.Caption = "Guest Checked Out"
                cmdDeposit.Enabled = False
                cmdAccept.Enabled = False
                If cmdAccept.Caption = "Save" Then cmdAccept.Enabled = False
        End Select
    End If
    DTArrival.Value = rs.Fields("Arrive_Date") & ""
    DTDeparture.Value = rs.Fields("Depart_Date") & ""
    cmbTitle.Text = rs.Fields("Title") & ""
    txtFName.Text = rs.Fields("First_Name") & ""
    txtLName.Text = rs.Fields("Last_Name") & ""
    txtFree.Text = Val(rs.Fields("Free_Nights") & "")
    txtID.Text = rs.Fields("ID_No") & ""
    cmbTitle.Text = rs.Fields("Title") & ""
    txtVehReg.Text = rs.Fields("Vehicle_No") & ""
    txtEmail.Text = rs.Fields("EMail") & ""
    txtTel.Text = rs.Fields("Tel_No") & ""
    txtMobile.Text = rs.Fields("Cell_No") & ""
    txtFax.Text = rs.Fields("Fax_No") & ""
    txtAddress.Text = rs.Fields("Address") & ""
    txtCode.Text = rs.Fields("Post_Code") & ""
    cmbCountry.Text = rs.Fields("Country") & ""
    cmbProvince = rs.Fields("Province") & ""
    cmbCity = rs.Fields("City") & ""
    txtContact.Text = rs.Fields("Contact_Person") & ""
    txtContactNo.Text = rs.Fields("Contact_no") & ""
    txtRemarks.Text = rs.Fields("Remarks") & ""
    cmbBusiness.Text = rs.Fields("Source") & ""
    cmbPay.Text = rs.Fields("Payment") & ""
    cmbBooked.Text = rs.Fields("Booked_By") & ""
    txtCredit.Text = rs.Fields("Credit_Card") & ""
    txtNights.Text = Val(DTDeparture.Value - DTArrival.Value)
    rs.Close
    Balance = 0
    ActiveReadServer "Select * from Room_Accounts where Res_No = '" & TillData.Res_No & "' order by Line_No"
    While Not rs.EOF
        Balance = Balance + (rs.Fields("Debit") - rs.Fields("Credit"))
        rs.MoveNext
    Wend
    rs.Close
    lblBal.Caption = Format(Balance, "0.00")
    lblTotRoom.Caption = Val(frmCheckin.txtNights.Text) - Val(frmCheckin.txtFree.Text) & " x Nights = " & Format(Val(frmCheckin.lblTotRate.Caption) * (Val(frmCheckin.txtNights.Text) - Val(frmCheckin.txtFree.Text)), "0.00")
fmType.Tag = ""
On Error GoTo 0
End Sub





Private Sub cmbBooked_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbBooked_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txt12to16.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub cmbBooked_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub cmbBusiness_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbBusiness_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                If picRates.Visible = False Then
                    cmbRate.SetFocus
                Else
                    txtContactNo.SetFocus
                End If
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub cmbCity_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbCity_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                cmbProvince.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub cmbCountry_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbCountry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtAddress.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub cmbPay_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbPay_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                cmbBusiness.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub cmbProvince_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbProvince_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                cmbCountry.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub cmbRate_Change()
    Calc_Rate
End Sub

Private Sub cmbRate_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtContactNo.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub cmbRoomNo_Change()
    With frmRes
        For i = 3 To .grdRes.Rows - 1
            If .grdRes.TextMatrix(i, 1) = cmbRoomNo.Text Then
                RowOld = .grdRes.Row
                ColOld = .grdRes.Col
                .grdRes.Row = i
                .grdRes.Cell(flexcpFontBold, RowOld, 0, RowOld, 1) = False
                .grdRes.Cell(flexcpFontBold, .grdRes.Row, 0, .grdRes.Row, 1) = True
                .grdRes.Cell(flexcpForeColor, RowOld, 0, RowOld, 1) = &H80000008
                .grdRes.Cell(flexcpForeColor, .grdRes.Row, 0, .grdRes.Row, 1) = &HC0&
                frmMain.stbBar.Panels(3).Text = Format(.grdRes.TextMatrix(2, .grdRes.Col), "dddd dd MMMM yyyy") & "   Row: " & .grdRes.Row & "    Col: " & .grdRes.Col
                txtDescription.Text = .grdRes.TextMatrix(.grdRes.Row, 0)
                Exit For
            End If
        Next i
    End With
End Sub
Private Sub cmbRoomNo_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbRoomNo_GotFocus()
    For i = 2 To 32
        If DTArrival.Value = frmRes.grdRes.TextMatrix(2, i) Then
            startcol = i
        End If
        If DTDeparture.Value = frmRes.grdRes.TextMatrix(2, i) Then
            stopcol = i
            Exit For
        End If
    Next i
    For i = 3 To frmRes.grdRes.Rows - 1
        For b = startcol To stopcol
                If frmRes.grdRes.TextMatrix(i, b) <> "" Then
                    If i <> frmRes.grdRes.Row Then
                        For ib = 0 To cmbRoomNo.ListCount - 1
                            If Val(cmbRoomNo.List(ib)) = frmRes.grdRes.TextMatrix(i, 1) Then
                                If startcol = b Then
                                    If InStr(frmRes.grdRes.TextMatrix(i, b), "D:") <> 0 And InStr(frmRes.grdRes.TextMatrix(i, b), "A:") = 0 Then Exit For
                                End If
                                If stopcol = b Then
                                    If InStr(frmRes.grdRes.TextMatrix(i, b), "D:") = 0 And InStr(frmRes.grdRes.TextMatrix(i, b), "A:") <> 0 Then Exit For
                                End If
                                cmbRoomNo.RemoveItem ib
                                Exit For
                               End If
                        Next ib
                    Exit For
                End If
            End If
        Next b
    Next i
End Sub
Private Sub cmbRoomNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtRemarks.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
End Sub
Private Sub cmbTitle_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
End Sub
Private Sub cmbTitle_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                DTDeparture.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub cmdAccept_Click()
    Select Case cmdAccept.Caption
        Case "Check In"
            ActiveReadServer "Select * from Reservations where Room_No = " & cmbRoomNo.Text & " and Res_Type = 2"
            If rs.RecordCount > 0 Then
                MsgBox rs.Fields("Title") & " " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name") & " is still Checked Into this Room." & Chr(13) & "Check the Guest out first before Checking In anyone else.", vbInformation, "HeroPOS"
                rs.Close
                Exit Sub
            End If
            rs.Close
            Response = MsgBox("Do you want to check in " & cmbTitle.Text & " " & txtFName.Text & " " & txtLName.Text & " now?", vbYesNo, "HeroPOS")
            If Response = vbYes Then
                ActiveUpdateServer "Update reservations set Res_Type = 2 where Res_No=" & lblResNo.Caption
                TillData.DocNo = 0
                TillData.Cashup_No = 0
                ActiveReadServer "Select * from Counters where User_no= " & UserRecord.User_Number & " and Finalized= 0"
                If rs.RecordCount > 0 Then
                    TillData.Cashup_No = rs.Fields("Cashup_No")
                Else
                    ActiveReadServer1 "Select isnull(max(Cashup_No),0)+1 as Cashup_No from Counters"
                    TillData.Cashup_No = rs1.Fields("Cashup_No")
                    rs1.Close
                    ActiveReadServer1 "Select Function_Key,Date_Time from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                    "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                    If rs1.RecordCount > 0 Then
                        If rs1.Fields("Function_Key") = 3 Then
                            ClockinTime = rs1.Fields("Date_Time") & ""
                        End If
                    End If
                    rs1.Close
                    ActiveUpdateServer "Insert into Counters (User_No,Cashup_No,Finalized,Counted,Shift_Start) values (" & UserRecord.User_Number & ", " & TillData.Cashup_No & ",0,0,'" & ClockinTime & "')"
                End If
                rs.Close
                
                Charge_Sales_Value = Val(lblTotRate.Caption) * (Val(txtNights.Text) - Val(txtFree.Text))
                
                ActiveUpdateServer "Update Counters set " & _
                "Charge_Sales_Value=isnull(Charge_Sales_Value,0) + " & Charge_Sales_Value & _
                ",Charge_Sales_Qty=isnull(Charge_Sales_Qty,0) + 1 " & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                ActiveUpdateServer "Update Counters set " & _
                "TaxableSales_Value = isnull(TaxableSales_Value,0) + " & Charge_Sales_Value & _
                ",TotalCalculatedTax_Value = isnull(TotalCalculatedTax_Value,0) + " & Charge_Sales_Value - (Charge_Sales_Value / 1.14) & _
                ",TotalCollectedTax_Value = isnull(TotalCollectedTax_Value,0) + " & Charge_Sales_Value - (Charge_Sales_Value / 1.14) & _
                " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                
                If TillData.DocNo = 0 Then
                    ActiveReadServer "Select (Select isnull(Max(Trans_No),0)+1 from Sales_Journal) as Trans_No,(Select isnull(Max(Invoice_No),0)+1 from Sales_Journal) as Invoice_No"
                    If rs.RecordCount > 0 Then
                        TillData.DocNo = rs.Fields("Invoice_No")
                        TillData.TransNo = rs.Fields("Trans_No")
                    End If
                    rs.Close
                End If
                
                ActiveReadServer "Select Location_No from Rooms where Room_No = " & cmbRoomNo.Text
                If rs.RecordCount > 0 Then
                    Location_No = rs.Fields("Location_No")
                End If
                rs.Close
                Function_Key = 20
                
                ActiveReadServer "Select * from Room_Rates where Rate_Type in (" & Replace(frmCheckin.lblRatestring.Caption, "-", ",") & ") order by Rate_Type"
                While Not rs.EOF
                    Qty = 1
                    QtySet = False
                    If Val(frmCheckin.txt0to5.Text) <> 0 And rs.Fields("Condition") = "Children under 5" Then
                        Qty = Val(frmCheckin.txt0to5.Text) * (Val(txtNights.Text) - Val(txtFree.Text))
                        QtySet = True
                    End If
                    If Val(frmCheckin.txt5.Text) <> 0 And rs.Fields("Condition") = "Children 5 to 12" Then
                        Qty = Val(frmCheckin.txt5.Text) * (Val(txtNights.Text) - Val(txtFree.Text))
                        QtySet = True
                    End If
                    If Val(frmCheckin.txt12to16.Text) <> 0 And rs.Fields("Condition") = "Children 12 to 15" Then
                        Qty = Val(frmCheckin.txt12to16.Text) * (Val(txtNights.Text) - Val(txtFree.Text))
                        QtySet = True
                    End If
                    If QtySet = False Then
                        Qty = Val(txtNights.Text) - Val(txtFree.Text)
                    End If
                    ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No, Conversion_Rate) values " & _
                    "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & Location_No & "," & Branch_No & "," & _
                    TillData.Cashup_No & ",'" & rs.Fields("Rate_Type") & "','0','" & Qty & "','0','14','1'," & rs.Fields("Room_Rate") * Qty & ",'" & rs.Fields("Description") & "','0','0','0','0',''," & cmbRoomNo.Text & "," & lblResNo.Caption & "," & Conversion_Rate & ")"
                    
                    rs.MoveNext
                Wend
                rs.Close
                Function_Key = 12
                
                Tax = Charge_Sales_Value - (Charge_Sales_Value / 1.14)
                
                ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No, Conversion_Rate) values " & _
                "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & Location_No & "," & Branch_No & "," & _
                TillData.Cashup_No & ",'0','0','0','0','" & Tax & "',''," & Charge_Sales_Value & ",'','0','0','0','0',''," & cmbRoomNo.Text & "," & lblResNo.Caption & "," & Conversion_Rate & ")"

                ActiveReadServer "Select Balance from Room_Accounts where Line_No in (Select max(Line_No) from Room_Accounts where Res_No =" & lblResNo.Caption & ")"
                If rs.RecordCount > 0 Then
                    OldBalance = Val(rs.Fields("Balance") & "")
                Else
                    OldBalance = 0
                End If
                rs.Close
                
                ActiveUpdateServer "INSERT INTO [Room_Accounts]([Transaction_Type],[Date_Time], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Res_No])" & _
                "VALUES('Accomodation',Getdate()," & TillData.DocNo & ",'" & TillData.Account_No & "'," & Charge_Sales_Value & ",0," & OldBalance + Charge_Sales_Value & "," & lblResNo.Caption & ")"
                
                If Trim(txtRemarks.Text) <> "" Then
                    ActiveReadServer1 "Select isnull(max(Task_No),0)+1 as Task_No from Room_Tasks"
                    Task_No = rs1.Fields("Task_No")
                    rs1.Close
                    ActiveUpdateServer "INSERT INTO [Room_Tasks]([Task_No], [Res_No], [Description], [Date_Time], [Remarks])" & _
                    " VALUES(" & Task_No & "," & lblResNo.Caption & ",'Check In Task',Getdate(),'" & txtRemarks.Text & "')"
                End If
                
                ActiveUpdateServer "Update Reservations set" & _
                " [Arr_Time] = '" & DTArrTime.Value & "'," & _
                " [Dep_Time] = '" & DTDepTime.Value & "',[Credit_Card] = '" & txtCredit.Text & "'," & _
                " [Res_No] = '" & lblResNo.Caption & "'," & _
                " [Room_No] = '" & cmbRoomNo.Text & "'," & _
                " [Free_Nights] = '" & txtFree.Text & "'," & _
                " [Workstation_No] = '" & Workstation_No & "'," & _
                " [Arrive_Date] = '" & DTArrival.Value & "'," & _
                " [Depart_Date] = '" & DTDeparture.Value & "'," & _
                " [Title] = '" & cmbTitle.Text & "'," & _
                " [First_Name] = '" & txtFName.Text & "'," & _
                " [Last_Name] = '" & txtLName.Text & "'," & _
                " [ID_No] = '" & txtID.Text & "'," & _
                " [Vehicle_No] = '" & txtVehReg.Text & "', [EMail] = '" & txtEmail.Text & "'," & _
                " [Tel_No] = '" & txtTel.Text & "', [Cell_No] = '" & txtMobile.Text & "'," & _
                " [Fax_No] = '" & txtFax.Text & "', [Address] = '" & txtAddress.Text & "'," & _
                " [Country] = '" & cmbCountry.Text & "', [Province] = '" & cmbProvince.Text & "'," & _
                " [City] = '" & cmbCity.Text & "'," & _
                " [Post_Code] = '" & Trim(txtCode.Text) & "', [Adults] = '" & txtAdults.Text & "'," & _
                " [Kid5to12] = '" & txt5.Text & "', [Kids12to16] = '" & txt12to16.Text & "', [Booked_By] = '" & cmbBooked.Text & "'," & _
                " [Contact_Person] = '" & txtContact.Text & "', [Contact_No] = '" & txtContactNo.Text & "'," & _
                " [Rate] = '" & Val(lblTotRate.Caption) & "', [Rate_Type] = '" & Val(Mid(cmbRate.Text, 1, InStr(cmbRate.Text, "-") - 1)) & "'," & _
                " [Source] = '" & cmbBusiness.Text & "', [Payment] = '" & cmbPay.Text & "'," & _
                " [Remarks] = '" & txtRemarks.Text & "'" & _
                " Where Res_No = " & lblResNo.Caption
                frmRes.Tag = "2"
                Unload Me
                Exit Sub
            End If
        Case "Check Out"
            Response = MsgBox("Do you want to check out " & cmbTitle.Text & " " & txtFName.Text & " " & txtLName.Text & " now?", vbYesNo, "HeroPOS")
            If Response = vbYes Then
                Balance = 0
                ActiveReadServer "Select * from Room_Accounts where Res_No = " & lblResNo.Caption
                While Not rs.EOF
                    Balance = Balance + Val(rs.Fields("Debit") - rs.Fields("Credit"))
                    rs.MoveNext
                Wend
                rs.Close
                If Balance > 0 Then
                    Response = MsgBox("There is still " & Format(Balance, "0.00") & " outstanding on this Account." & Chr(13) & "You have to Settle the Account before Checking Out." & Chr(13) & "Do you want to Charge this to a Debtor?", vbYesNo, "HeroPOS")
                    If Response = vbYes Then
                        Load frmCharge
                        frmCharge.cmdInput(1).Enabled = False
                        frmCharge.cmdInput(2).Enabled = False
                        frmCharge.cmdInput(3).Enabled = False
                        frmCharge.lblHeading.Tag = "Select a Debtor to Close the Room to."
                        frmCharge.Show vbModal
                    End If
                Else
                    ActiveUpdateServer "Update Reservations set Res_Type = 3 where Res_No=" & lblResNo.Caption
                End If
                frmRes.Tag = "3"
                Unload Me
                Exit Sub
            End If
        Case "Accept", "Save"
            Res_Type = 0
            DTDeparture.Tag = ""
            ActiveReadServer "Select * from Reservations where Res_No = " & Val(lblResNo.Caption)
            If rs.RecordCount > 0 Then
                Res_Type = rs.Fields("Res_Type")
                
                
                If Val(Format(rs.Fields("Depart_Date"), vbGeneralDate)) - Val(Format(rs.Fields("Arrive_Date"), vbGeneralDate)) < Val(txtNights.Text) And rs.Fields("Res_Type") = 2 Then
                 '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                    
                    
                    
                    Response = MsgBox("Do you want to Charge the Guest for the Extra Nights?" & Chr(13) & "Departure Date Moved to " & Format(DTDeparture.Value, "dddd dd MMM yyyy"), vbYesNo, "HeroPOS")
                    If Response = vbYes Then
                        ChargeNights = Val(txtNights.Text) - (Val(Format(rs.Fields("Depart_Date"), vbGeneralDate)) - Val(Format(rs.Fields("Arrive_Date"), vbGeneralDate)))
                        DTDeparture.Tag = "1"
                        rs.Close
                        TillData.DocNo = 0
                        TillData.Cashup_No = 0
                        ActiveReadServer "Select * from Counters where User_no= " & UserRecord.User_Number & " and Finalized= 0"
                        If rs.RecordCount > 0 Then
                            TillData.Cashup_No = rs.Fields("Cashup_No")
                        Else
                        
                        
                            ActiveReadServer1 "Select isnull(max(Cashup_No),0)+1 as Cashup_No from Counters"
                            TillData.Cashup_No = rs1.Fields("Cashup_No")
                            rs1.Close
                            ActiveReadServer1 "Select Function_Key,Date_Time from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                            "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                            If rs1.RecordCount > 0 Then
                                If rs1.Fields("Function_Key") = 3 Then
                                    ClockinTime = rs1.Fields("Date_Time") & ""
                                End If
                            End If
                            rs1.Close
                            ActiveUpdateServer "Insert into Counters (User_No,Cashup_No,Finalized,Counted,Shift_Start) values (" & UserRecord.User_Number & ", " & TillData.Cashup_No & ",0,0,'" & ClockinTime & "')"
                                 
                     
                        End If
                        rs.Close
                        
                        Charge_Sales_Value = Val(lblTotRate.Caption) * Val(ChargeNights)
                        ActiveUpdateServer "Update Counters set " & _
                        "Charge_Sales_Value=isnull(Charge_Sales_Value,0) + " & Charge_Sales_Value & _
                        ",Charge_Sales_Qty=isnull(Charge_Sales_Qty,0) + 1 " & _
                        " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                        ActiveUpdateServer "Update Counters set " & _
                        "TaxableSales_Value = isnull(TaxableSales_Value,0) + " & Charge_Sales_Value & _
                        ",TotalCalculatedTax_Value = isnull(TotalCalculatedTax_Value,0) + " & Charge_Sales_Value - (Charge_Sales_Value / 1.14) & _
                        ",TotalCollectedTax_Value = isnull(TotalCollectedTax_Value,0) + " & Charge_Sales_Value - (Charge_Sales_Value / 1.14) & _
                        " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                        
                        If TillData.DocNo = 0 Then
                            ActiveReadServer "Select (Select isnull(Max(Trans_No),0)+1 from Sales_Journal) as Trans_No,(Select isnull(Max(Invoice_No),0)+1 from Sales_Journal) as Invoice_No"
                            If rs.RecordCount > 0 Then
                                TillData.DocNo = rs.Fields("Invoice_No")
                                TillData.TransNo = rs.Fields("Trans_No")
                            End If
                            rs.Close
                        End If
                        
                        ActiveReadServer "Select Location_No from Rooms where Room_No = " & cmbRoomNo.Text
                        If rs.RecordCount > 0 Then
                            Location_No = rs.Fields("Location_No")
                        End If
                        rs.Close
        '????????????????????????????????
                        Function_Key = 20
                        ActiveReadServer "Select * from Room_Rates where Rate_Type in (" & Replace(frmCheckin.lblRatestring.Caption, "-", ",") & ") order by Rate_Type"
                        While Not rs.EOF
                            Qty = 1
                            QtySet = False
                            If Val(frmCheckin.txt0to5.Text) <> 0 And rs.Fields("Condition") = "Children under 5" Then
                                Qty = Val(frmCheckin.txt0to5.Text) * Val(ChargeNights)
                                QtySet = True
                            End If
                            If Val(frmCheckin.txt5.Text) <> 0 And rs.Fields("Condition") = "Children 5 to 12" Then
                                Qty = Val(frmCheckin.txt5.Text) * Val(ChargeNights)
                                QtySet = True
                            End If
                            If Val(frmCheckin.txt12to16.Text) <> 0 And rs.Fields("Condition") = "Children 12 to 15" Then
                                Qty = Val(frmCheckin.txt12to16.Text) * Val(ChargeNights)
                                QtySet = True
                            End If
                            If QtySet = False Then
                                Qty = Val(ChargeNights)
                            End If
                            ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No, Conversion_Rate) values " & _
                            "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & Location_No & "," & Branch_No & "," & _
                            TillData.Cashup_No & ",'" & rs.Fields("Rate_Type") & "','0','" & Qty & "','0','14','1'," & rs.Fields("Room_Rate") * Qty & ",'" & rs.Fields("Description") & "','0','0','0','0',''," & cmbRoomNo.Text & "," & lblResNo.Caption & "," & Conversion_Rate & ")"
                            
                            rs.MoveNext
                        Wend
                        rs.Close
                        Function_Key = 12
                                                Tax = Charge_Sales_Value - (Charge_Sales_Value / 1.14)
                        
                        ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No, Conversion_Rate) values " & _
                        "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & Location_No & "," & Branch_No & "," & _
                        TillData.Cashup_No & ",'0','0','0','0','" & Tax & "',''," & Charge_Sales_Value & ",'','0','0','0','0',''," & cmbRoomNo.Text & "," & lblResNo.Caption & "," & Conversion_Rate & ")"
                        
                        ActiveReadServer "Select Balance from Room_Accounts where Line_No in (Select max(Line_No) from Room_Accounts where Res_No =" & lblResNo.Caption & ")"
                        If rs.RecordCount > 0 Then
                            OldBalance = Val(rs.Fields("Balance") & "")
                        Else
                            OldBalance = 0
                        End If
                        rs.Close
                        
                        '????????????????????????
                        
                        ActiveUpdateServer "INSERT INTO [Room_Accounts]([Transaction_Type],[Date_Time], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Res_No])" & _
                        "VALUES('Accomodation',Getdate()," & TillData.DocNo & ",'" & TillData.Account_No & "'," & Charge_Sales_Value & ",0," & OldBalance + Charge_Sales_Value & "," & lblResNo.Caption & ")"
                    End If
                    GoTo Dones
                End If '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                
               
                
                
                
                
                
                 ActiveReadServer "Select * from Reservations where Res_No = " & Val(lblResNo.Caption)
            If rs.RecordCount > 0 Then
                Res_Type = rs.Fields("Res_Type")
                        '**********************
                         If Val(Format(rs.Fields("Depart_Date"), vbGeneralDate)) - Val(Format(rs.Fields("Arrive_Date"), vbGeneralDate)) > Val(txtNights.Text) And rs.Fields("Res_Type") = 2 Then
                      Response = MsgBox("Do you want to subtract Charge for less Nights?" & Chr(13) & "Departure Date Moved to " & Format(DTDeparture.Value, "dddd dd MMM yyyy"), vbYesNo, "HeroPOS")
                        If Response = vbYes Then 'OOOOOOOOOOOOOOOOOOOOOO
                        ChargeNights = Val(txtNights.Text) - (Val(Format(rs.Fields("Depart_Date"), vbGeneralDate)) - Val(Format(rs.Fields("Arrive_Date"), vbGeneralDate)))
                        DTDeparture.Tag = "1"
                        rs.Close
                         ActiveReadServer "Select * from Counters where User_no= " & UserRecord.User_Number & " and Finalized= 0"
                        If rs.RecordCount > 0 Then 'UUUUUUUUUUUUUUUUUUUUUUUUUUU
                           TillData.Cashup_No = rs.Fields("Cashup_No")
                        Else 'UUUUUUUUUUUUUUUUUUUUUUUU
                            ActiveReadServer1 "Select isnull(max(Cashup_No),0)+1 as Cashup_No from Counters"
                            TillData.Cashup_No = rs1.Fields("Cashup_No")
                            rs1.Close
                            ActiveReadServer1 "Select Function_Key,Date_Time from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
                            "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
                            If rs1.RecordCount > 0 Then
                                If rs1.Fields("Function_Key") = 3 Then
                                    ClockinTime = rs1.Fields("Date_Time") & ""
                                End If
                            End If
                            End If 'UUUUUUUUUUUUUUUUUUUUUUUUUU
                            rs.Close
                           'ActiveUpdateServer "Insert into Counters (User_No,Cashup_No,Finalized,Counted,Shift_Start) values (" & UserRecord.User_Number & ", " & TillData.Cashup_No & ",0,0,'" & ClockinTime & "')"


                        Charge_Sales_Value = Val(lblTotRate.Caption) * Val(ChargeNights)
                        ActiveUpdateServer "Update Counters set " & _
                        "Charge_Sales_Value=isnull(Charge_Sales_Value,0) + " & Charge_Sales_Value & _
                        ",Charge_Sales_Qty=isnull(Charge_Sales_Qty,0) + 1 " & _
                        " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
                        ActiveUpdateServer "Update Counters set " & _
                        "TaxableSales_Value = isnull(TaxableSales_Value,0) + " & Charge_Sales_Value & _
                        ",TotalCalculatedTax_Value = isnull(TotalCalculatedTax_Value,0) + " & Charge_Sales_Value - (Charge_Sales_Value / 1.14) & _
                        ",TotalCollectedTax_Value = isnull(TotalCollectedTax_Value,0) + " & Charge_Sales_Value - (Charge_Sales_Value / 1.14) & _
                        " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number

                        ActiveReadServer "Select Balance from Room_Accounts where Line_No in (Select max(Line_No) from Room_Accounts where Res_No =" & lblResNo.Caption & ")"
                        If rs.RecordCount > 0 Then
                            OldBalance = Val(rs.Fields("Balance") & "")
                        Else
                            OldBalance = 0
                        End If
                        rs.Close
                        
                        
                         If TillData.DocNo = 0 Then
                            ActiveReadServer "Select (Select isnull(Max(Trans_No),0)+1 from Sales_Journal) as Trans_No,(Select isnull(Max(Invoice_No),0)+1 from Sales_Journal) as Invoice_No"
                            If rs.RecordCount > 0 Then
                                TillData.DocNo = rs.Fields("Invoice_No")
                                TillData.TransNo = rs.Fields("Trans_No")
                            End If
                            rs.Close
                        End If
                        
                        ActiveReadServer "Select Location_No from Rooms where Room_No = " & cmbRoomNo.Text
                        If rs.RecordCount > 0 Then
                            Location_No = rs.Fields("Location_No")
                        End If
                        rs.Close
        
                        Function_Key = 20
                         '?????????????????????????
                         Function_Key = 20
                        ActiveReadServer "Select * from Room_Rates where Rate_Type in (" & Replace(frmCheckin.lblRatestring.Caption, "-", ",") & ") order by Rate_Type"
                        While Not rs.EOF
                            Qty = 1
                            QtySet = False
                            If Val(frmCheckin.txt0to5.Text) <> 0 And rs.Fields("Condition") = "Children under 5" Then
                                Qty = Val(frmCheckin.txt0to5.Text) * Val(ChargeNights)
                                QtySet = True
                            End If
                            If Val(frmCheckin.txt5.Text) <> 0 And rs.Fields("Condition") = "Children 5 to 12" Then
                                Qty = Val(frmCheckin.txt5.Text) * Val(ChargeNights)
                                QtySet = True
                            End If
                            If Val(frmCheckin.txt12to16.Text) <> 0 And rs.Fields("Condition") = "Children 12 to 15" Then
                                Qty = Val(frmCheckin.txt12to16.Text) * Val(ChargeNights)
                                QtySet = True
                            End If
                            If QtySet = False Then
                                Qty = Val(ChargeNights)
                            End If
                            ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No,Conversion_Rate) values " & _
                            "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & Location_No & "," & Branch_No & "," & _
                            TillData.Cashup_No & ",'" & rs.Fields("Rate_Type") & "','0','" & Qty & "','0','14','1'," & rs.Fields("Room_Rate") * Qty & ",'" & rs.Fields("Description") & "','0','0','0','0',''," & cmbRoomNo.Text & "," & lblResNo.Caption & "," & Conversion_Rate & ")"
                            
                            rs.MoveNext
                        Wend
                        rs.Close
                        Function_Key = 12
                                                Tax = Charge_Sales_Value - (Charge_Sales_Value / 1.14)
                        
                        ActiveUpdateServer1 "Insert into Sales_Journal (Date_Time,Workstation_No,User_No,Trans_No,Invoice_No,Function_key,Location,Branch_No,Cashup_No,Product_Code,Department_No,Qty,Ave_Cost,Sales_Tax,Tax_Type,Line_Total,Extra,Table_No,Tab_No,Covers,User_Overide,Account_No,Room_No,Res_No,Conversion_Rate) values " & _
                        "(GetDate()," & Workstation_No & "," & UserRecord.User_Number & "," & TillData.TransNo & "," & TillData.DocNo & "," & Val(Function_Key) & "," & Location_No & "," & Branch_No & "," & _
                        TillData.Cashup_No & ",'0','0','0','0','" & Tax & "',''," & Charge_Sales_Value & ",'','0','0','0','0',''," & cmbRoomNo.Text & "," & lblResNo.Caption & "," & Conversion_Rate & ")"
                        
                        ActiveReadServer "Select Balance from Room_Accounts where Line_No in (Select max(Line_No) from Room_Accounts where Res_No =" & lblResNo.Caption & ")"
                        If rs.RecordCount > 0 Then
                            OldBalance = Val(rs.Fields("Balance") & "")
                        Else
                            OldBalance = 0
                        End If
                        rs.Close
                         '????????????????????????
                        ActiveUpdateServer "INSERT INTO [Room_Accounts]([Transaction_Type],[Date_Time], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Res_No])" & _
                        "VALUES('Accomodation',Getdate()," & TillData.DocNo & ",'" & TillData.Account_No & "'," & Charge_Sales_Value & ",0," & OldBalance + Charge_Sales_Value & "," & lblResNo.Caption & ")"




''
                       End If 'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
''                        rs.Close
                End If '**********************
                End If
                
                
                
                
                
                
Dones:
                
                
                txtNights.Text = Val(DTDeparture.Value - DTArrival.Value) - 1
                ActiveUpdateServer "Update Reservations set" & _
                " [Arr_Time] = '" & DTArrTime.Value & "'," & _
                " [Dep_Time] = '" & DTDepTime.Value & "',[Credit_Card] = '" & txtCredit.Text & "'," & _
                " [Res_No] = '" & lblResNo.Caption & "'," & _
                " [Room_No] = '" & cmbRoomNo.Text & "'," & _
                " [Free_Nights] = '" & txtFree.Text & "'," & _
                " [Workstation_No] = '" & Workstation_No & "'," & _
                " [Arrive_Date] = '" & DTArrival.Value & "'," & _
                " [Depart_Date] = '" & DTDeparture.Value & "'," & _
                " [Title] = '" & cmbTitle.Text & "'," & _
                " [First_Name] = '" & txtFName.Text & "'," & _
                " [Last_Name] = '" & txtLName.Text & "'," & _
                " [ID_No] = '" & txtID.Text & "'," & _
                " [Vehicle_No] = '" & txtVehReg.Text & "', [EMail] = '" & txtEmail.Text & "'," & _
                " [Tel_No] = '" & txtTel.Text & "', [Cell_No] = '" & txtMobile.Text & "'," & _
                " [Fax_No] = '" & txtFax.Text & "', [Address] = '" & txtAddress.Text & "'," & _
                " [Country] = '" & cmbCountry.Text & "', [Province] = '" & cmbProvince.Text & "'," & _
                " [City] = '" & cmbCity.Text & "'," & _
                " [Post_Code] = '" & Trim(txtCode.Text) & "', [Adults] = '" & txtAdults.Text & "'," & _
                " [Kid5to12] = '" & txt5.Text & "', [Kids12to16] = '" & txt12to16.Text & "', [Booked_By] = '" & cmbBooked.Text & "'," & _
                " [Contact_Person] = '" & txtContact.Text & "', [Contact_No] = '" & txtContactNo.Text & "'," & _
                " [Rate] = '" & Val(lblTotRate.Caption) & "', [Rate_Type] = '" & Val(Mid(cmbRate.Text, 1, InStr(cmbRate.Text, "-") - 1)) & "'," & _
                " [Source] = '" & cmbBusiness.Text & "', [Payment] = '" & cmbPay.Text & "'," & _
                " [Remarks] = '" & txtRemarks.Text & "', [Res_Type] =  '" & Res_Type & "'" & _
                " Where Res_No = " & lblResNo.Caption
            Else
                ActiveReadServer1 "Select isnull(max(Res_No),0)+1 as Res_No from Reservations"
                lblResNo.Caption = rs1.Fields("Res_No")
                rs1.Close
                ActiveUpdateServer "INSERT INTO [Reservations]([Free_Nights],[Arr_Time],[Dep_Time],[Res_No], [Room_No], [Workstation_No], [Arrive_Date], [Depart_Date], [Title], [First_Name], [Last_Name], [ID_No], [Vehicle_No], [EMail], [Tel_No], [Cell_No], [Fax_No],[Address], [Country], [Province], [City], [Post_Code], [Adults], [Kid5to12], [Kids12to16], [Booked_By], [Contact_Person], [Contact_No], [Rate], [Rate_Type], [Source], [Payment], [Remarks],[Res_Type],[Credit_Card])" & _
                " VALUES('" & txtFree.Text & "','" & DTArrTime.Value & "','" & DTDepTime.Value & "','" & lblResNo.Caption & "','" & cmbRoomNo.Text & "','" & Workstation_No & "','" & DTArrival.Value & "','" & DTDeparture.Value & "','" & cmbTitle.Text & "','" & txtFName.Text & "','" & txtLName.Text & "','" & txtID.Text & "','" & txtVehReg.Text & "','" & txtEmail.Text & "','" & txtTel.Text & "','" & txtMobile.Text & "','" & txtFax.Text & "','" & txtAddress.Text & "','" & cmbCountry.Text & "','" & cmbProvince.Text & "','" & cmbCity.Text & "','" & txtCode.Text & "','" & txtAdults.Text & "','" & txt5.Text & "','" & txt12to16.Text & "','" & cmbBooked.Text & "','" & txtContact.Text & "','" & txtContactNo.Text & "','" & Val(lblTotRate.Caption) & "','" & Val(Mid(cmbRate.Text, 1, InStr(cmbRate.Text, "-") - 1)) & "','" & cmbBusiness.Text & "','" & cmbPay.Text & "','" & txtRemarks.Text & "',0,'" & txtCredit.Text & "')"
            End If
            
            If CheckBox1.Value <> 0 Then
            If txtID.Text <> "" Then
            ActiveReadServer "Select * from Adressbook where ID_No = '" & txtID.Text & "'"
            If rs.RecordCount > 0 Then
            MsgBox "Id already in database could not save information in Addressbook", vbOKOnly
            
            If rs.RecordCount = 0 Then
               ActiveUpdateServer "INSERT INTO Adressbook([ID_No],[Title],[First_Name],[Last_Name],[Tel_No], [Fax_No], [Cell_No], [Address],[City],[Post_Code], [Province], [Country],[Remarks],[Email],[Last_Booking])" & _
               "VALUES('" & txtID.Text & "', '" & cmbTitle.Text & "' , '" & txtFName.Text & "','" & txtLName.Text & "','" & txtTel.Text & "','" & txtFax.Text & "', '" & txtMobile.Text & "','" & txtAddress.Text & "','" & cmbCity.Text & "','" & Trim(txtCode.Text) & "','" & cmbProvince.Text & "','" & cmbCountry.Text & "','" & txtRemarks.Text & "','" & txtEmail.Text & "','" & DTArrTime.Value & " ')"
                        
                    
'            Else
'
'                ActiveUpdateServer "Update Adressbook set" & _
'                " [ID_No] = '" & txtId.Text & "'," & _
'                " [Title] = '" & cmbTitle.Text & "'," & _
'                " [First_Name] = '" & txtFName.Text & "'," & _
'                " [Last_Name] = '" & txtLName.Text & "'," & _
'                " [Tel_No] = '" & txttel.Text & "'," & _
'                " [Fax_No] = '" & txtFax.Text & "'," & _
'                " [Cell_No] = '" & txtMobile.Text & "'," & _
'                " [Address] = '" & txtAddress.Text & "'," & _
'                " [City] = '" & cmbCity.Text & "'," & _
'                " [Post_Code] = '" & Trim(txtCode.Text) & "'," & _
'                " [Province] = '" & cmbProvince.Text & "'," & _
'                " [Country] = '" & cmbCountry.Text & "'," & _
'                " [Remarks] = '" & txtRemarks.Text & "'," & _
'                " [Email] = '" & txtEmail.Text & "'," & _
'                " [Last_Booking] = '" & DTArrTime.Value & "'" & _
'                " Where [ID_No] = '" & txtId.Text & "'"
'
                End If
                End If
                End If
            End If
            
            
    End Select
    If DTDeparture.Tag <> "" Then
        frmRes.Tag = "4"
    Else
        frmRes.Tag = "1"
    End If
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    frmRes.Tag = ""
    Unload Me
End Sub

Private Sub cmdCondition_Click()
    frmConditions.Show vbModal
End Sub
Private Sub cmdDeposit_Click()
    Load frmCharge
    frmCharge.picReceive.Visible = True
    frmCharge.cmdInput(0).Enabled = False
    frmCharge.cmdInput(1).Enabled = False
    frmCharge.cmdInput(2).Enabled = False
    frmCharge.cmdInput(3).Value = 1
    frmCharge.lblHeading.Caption = "Please Enter Deposit for Room: " & cmbRoomNo.Text
    frmCharge.grdMain.Rows = 1
    ActiveReadServer "Select * from Room_Accounts where Res_No = " & lblResNo.Caption & " order by Line_No"
    Balance = 0
    While Not rs.EOF
        frmCharge.grdMain.Rows = frmCharge.grdMain.Rows + 1
        frmCharge.grdMain.Row = frmCharge.grdMain.Rows - 1
        If rs.Fields("Transaction_Type") & "" = "Invoice" Then
            frmCharge.grdMain.TextMatrix(frmCharge.grdMain.Row, 0) = "Sales Invoice"
            frmCharge.grdMain.TextMatrix(frmCharge.grdMain.Row, 1) = Format(rs.Fields("Debit") & "", "0.00")
            Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
        End If
        If rs.Fields("Transaction_Type") & "" = "Receipt" Then
            frmCharge.grdMain.TextMatrix(frmCharge.grdMain.Row, 0) = "Receive on Account"
            frmCharge.grdMain.TextMatrix(frmCharge.grdMain.Row, 1) = Format(rs.Fields("Credit") & "", "0.00")
            Balance = Balance - Format(rs.Fields("Credit") & "", "0.00")
        End If
        If rs.Fields("Transaction_Type") & "" = "Deposit" Then
            frmCharge.grdMain.TextMatrix(frmCharge.grdMain.Row, 0) = "Deposit Received"
            frmCharge.grdMain.TextMatrix(frmCharge.grdMain.Row, 1) = Format(rs.Fields("Credit") & "", "0.00")
            Balance = Balance - Format(rs.Fields("Credit") & "", "0.00")
        End If
        rs.MoveNext
    Wend
    frmCharge.grdMain.ShowCell frmCharge.grdMain.Row, 0
    rs.Close
    frmCharge.lblTender.Caption = Format(Balance, "0.00")
    frmCharge.Show vbModal
    If frmCharge.Tag = "" Or frmCharge.Tag = "Cancel" Then
        Unload frmCharge
       Exit Sub
    Else
        Deposit = frmCharge.lblHeading.Caption
        ActiveReadServer1 "Select isnull(max(Invoice_No),0)+1 as Deposit_No from Room_Accounts where Transaction_Type = 'Deposit'"
        Deposit_No = rs1.Fields("Deposit_No")
        rs1.Close
        
        Balance = 0
        ActiveReadServer "Select Balance from Room_Accounts where Res_No = " & lblResNo.Caption & " order by Line_No"
        If rs.RecordCount > 0 Then
            rs.MoveLast
            Balance = rs.Fields("Balance")
        End If
        rs.Close
        
        Tender_Type = Mid(frmCharge.Tag, InStr(frmCharge.Tag, "|") + 1)
        
        ActiveUpdateServer "INSERT INTO [Room_Accounts]([User_No],[Date_Time],[Transaction_Type], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Res_No],[Tender_Type])" & _
        "VALUES(" & UserRecord.User_Number & ",Getdate(),'Deposit'," & Deposit_No & "," & cmbRoomNo.Text & ",0," & Deposit & "," & Balance + (Deposit * -1) & "," & lblResNo.Caption & ",'" & Tender_Type & " ')"
        DoEvents
        
        If fmType.Caption = "Provisional Booking" Then
            ActiveUpdateServer "Update Reservations set Res_Type = 1 where Res_No = " & lblResNo.Caption
            LoadResDetail lblResNo.Caption
        End If
        DoEvents
        
        TillData.Cashup_No = 0
        ActiveReadServer "Select * from Counters where User_no= " & UserRecord.User_Number & " and Finalized= 0"
        If rs.RecordCount > 0 Then
            TillData.Cashup_No = rs.Fields("Cashup_No")
        Else
            ActiveReadServer1 "Select isnull(max(Cashup_No),0)+1 as Cashup_No from Counters"
            TillData.Cashup_No = rs1.Fields("Cashup_No")
            rs1.Close
            ActiveReadServer1 "Select Function_Key,Date_Time from User_Journal where user_No= " & UserRecord.User_Number & " and line_No = " & _
            "(Select Max(Line_No) from User_Journal where function_Key in (3,4) and User_No=" & UserRecord.User_Number & ")"
            If rs1.RecordCount > 0 Then
                If rs1.Fields("Function_Key") = 3 Then
                    ClockinTime = rs1.Fields("Date_Time") & ""
                End If
            End If
            rs1.Close
            ActiveUpdateServer "Insert into Counters (User_No,Cashup_No,Finalized,Counted,Shift_Start) values (" & UserRecord.User_Number & ", " & TillData.Cashup_No & ",0,0,'" & ClockinTime & "')"
        End If
        rs.Close

        ActiveUpdateServer "Update Counters set " & _
        "Loans_Value = isnull(Loans_Value,0) + " & (Deposit) & _
        ",Loans_Qty=isnull(Loans_Qty,0) +" & 1 & _
        " where Cashup_No=" & TillData.Cashup_No & " And User_No = " & UserRecord.User_Number
        DoEvents
        SaveRow = frmRes.grdRes.Row
        savecol = frmRes.grdRes.Col
        frmRes.LoadRes
        
        Print_Receive_on_Account lblResNo.Caption, cmbRoomNo.Text, "", Deposit, 0, Deposit_No, Tender_Type
        
        Response = MsgBox("Do you want to Print the Check In Form now?", vbYesNo, "HeroPOS")
        If Response = vbYes Then
            frmRes.grdRes.Row = SaveRow
            frmRes.grdRes.Col = savecol
            TillData.Res_No = Val(Mid(frmRes.grdRes.TextMatrix(frmRes.grdRes.Row, frmRes.grdRes.Col), InStr(frmRes.grdRes.TextMatrix(frmRes.grdRes.Row, frmRes.grdRes.Col), ":") + 1, (InStr(frmRes.grdRes.TextMatrix(frmRes.grdRes.Row, frmRes.grdRes.Col), "-") - 1) - (InStr(frmRes.grdRes.TextMatrix(frmRes.grdRes.Row, frmRes.grdRes.Col), ":") + 1)))
            DoEvents
            Unload frmCharge
            DoEvents
            Unload frmCheckin
            rptCheckin.Show
            Exit Sub
        End If
    End If
    DoEvents
    Unload frmCharge
End Sub

Private Sub DTArrival_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    txtNights.Text = Val(DTDeparture.Value - DTArrival.Value)
    Calc_Rate
End Sub
Private Sub DTDeparture_Change()
    With frmRes
        For i = 3 To .grdRes.Rows - 1
            If .grdRes.TextMatrix(i, 1) = cmbRoomNo.Text Then
                .grdRes.Row = i
                Exit For
            End If
        Next i
        For i = .grdRes.Col + 1 To .grdRes.Cols - 2
            If DTDeparture.Value = .grdRes.TextMatrix(2, i) And .grdRes.TextMatrix(.grdRes.Row, i) = "" Then
                Exit For
            End If
            If .grdRes.TextMatrix(.grdRes.Row, i) <> "" Then
                If TillData.Res_No <> Val(Mid(.grdRes.TextMatrix(.grdRes.Row, i), InStr(.grdRes.TextMatrix(.grdRes.Row, i), ":") + 1, (InStr(.grdRes.TextMatrix(.grdRes.Row, i), "-") - 1) - (InStr(.grdRes.TextMatrix(.grdRes.Row, i), ":") + 1))) Then
                    If Format(DTDeparture.Value, vbGeneralDate) > Format(.grdRes.TextMatrix(2, i), vbGeneralDate) Then
                        DTDeparture.Value = .grdRes.TextMatrix(2, i)
                    End If
                    GoTo far
                Else
                    If DTDeparture.Value = .grdRes.TextMatrix(2, i) Then Exit For
                End If
            End If
        Next i
        If Format(DTDeparture.Value, vbGeneralDate) > Format(.grdRes.TextMatrix(2, 32), vbGeneralDate) Then
            ActiveReadServer "Select Arrive_Date from Reservations where Room_No = " & cmbRoomNo.Text & " and (Arrive_Date < '" & DTDeparture.Value & "' and Arrive_Date > '" & DTArrival.Value & "')"
            If rs.RecordCount > 0 Then
                DTDeparture.Value = rs.Fields("Arrive_Date")
            End If
            rs.Close
        End If
far:
        If Format(DTDeparture.Value, vbGeneralDate) < Format(DTArrival.Value, vbGeneralDate) Or Format(DTDeparture.Value, vbGeneralDate) = Format(DTArrival.Value, vbGeneralDate) Then
            DTDeparture.Value = DateAdd("D", 1, DTArrival.Value)
        End If
        txtNights.Text = Val(DTDeparture.Value - DTArrival.Value)
        Calc_Rate
    End With
End Sub

Private Sub Form_Activate()
If txtFName.Tag <> "" Then Exit Sub
On Error Resume Next
    txtFName.SetFocus
On Error GoTo 0
End Sub

Private Sub Form_Load()
CheckBox1.Value = 1
    With frmRes
        DTDeparture.Tag = ""
        cmbBusiness.Clear
        cmbBusiness.AddItem "Guests"
        cmbBusiness.AddItem "Travel Agent"
        cmbBusiness.AddItem "Advertisment"
        cmbBusiness.Text = "Guests"
        cmbBooked.Clear
        cmbBooked.AddItem "Guests"
        cmbBooked.AddItem "Travel Agent"
        cmbBooked.Text = "Guests"
        cmbPay.Clear
        cmbPay.AddItem "Guests"
        cmbPay.AddItem "Travel Agent"
        cmbPay.Text = "Guests"
        cmbRate.Clear
        ActiveReadServer1 "Select * from Room_Rates  where Active=1 order by Rate_Type"
        While Not rs1.EOF
            cmbRate.AddItem rs1.Fields("Rate_Type") & " - " & rs1.Fields("Description") & " > " & Format(rs1.Fields("Room_Rate"), "0.00")
            rs1.MoveNext
        Wend
        rs1.MoveFirst
        cmbRate.Text = rs1.Fields("Rate_Type") & " - " & rs1.Fields("Description") & " > " & Format(rs1.Fields("Room_Rate"), "0.00")
        rs1.Close
        ActiveReadServer "Select Room_Rate from Rooms where Room_No = " & frmRes.grdRes.ValueMatrix(frmRes.grdRes.Row, 1)
        If rs.RecordCount > 0 Then
            For i = 0 To cmbRate.ListCount - 1
                If Val(Mid(cmbRate.List(i, 0), 1, InStr(cmbRate.List(i, 0), "-") - 1)) = Val(rs.Fields("Room_Rate") & "") Then
                    cmbRate.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        rs.Close
        cmbCountry.Clear
        cmbCountry.AddItem "Angola"
        cmbCountry.AddItem "Botswana"
        cmbCountry.AddItem "Malawi"
        cmbCountry.AddItem "Mosambique"
        cmbCountry.AddItem "Namibia"
        cmbCountry.AddItem "South Africa"
        cmbCountry.AddItem "Zimbabwe"
        cmbCountry.Text = "South Africa"
        cmbCity.Clear
        cmbCity.AddItem "Bloemfontein"
        cmbCity.AddItem "Cape Town"
        cmbCity.AddItem "Durban"
        cmbCity.AddItem "East London"
        cmbCity.AddItem "Johannesburg"
        cmbCity.AddItem "Port Elizabeth"
        cmbCity.AddItem "Pretoria"
        cmbCity.AddItem "Polokwane"
        cmbCity.AddItem "Windhoek"
        cmbCity.AddItem "Gaborone"
        cmbCity.AddItem "Maputo"
        cmbTitle.Clear
        cmbTitle.AddItem "Dr."
        cmbTitle.AddItem "Prof."
        cmbTitle.AddItem "Miss."
        cmbTitle.AddItem "Mr."
        cmbTitle.AddItem "Mrs."
        cmbTitle.AddItem "Ms."
        cmbTitle.Text = "Mr."
        cmbProvince.Clear
        ActiveReadServer "Select * from Regions"
        While Not rs.EOF
            cmbProvince.AddItem rs.Fields("Region_Name")
            rs.MoveNext
        Wend
        rs.Close
        cmbRoomNo.Clear
        For i = 3 To .grdRes.Rows - 1
            cmbRoomNo.AddItem .grdRes.TextMatrix(i, 1)
        Next i
        For i = 0 To frmRes.cmdTop.Count - 1
            If frmRes.cmdTop(i).Value = Down Then
                If frmRes.cmdTop(i).Caption = "Check In/Out" Then
                    cmdAccept.Caption = ""
                    Exit For
                Else
                    cmdAccept.Caption = "Accept"
                End If
                If frmRes.cmdTop(i).Caption = "Change" Then
                    cmdAccept.Caption = "Save"
                    Exit For
                Else
                    cmdAccept.Caption = "Accept"
                End If
            End If
        Next i
        cmbRoomNo.Text = .grdRes.TextMatrix(.grdRes.Row, 1)
        txtDescription.Text = .grdRes.TextMatrix(.grdRes.Row, 0)
        If .grdRes.TextMatrix(.grdRes.Row, .grdRes.Col) <> "" Then
            lblUser.Caption = UserRecord.User_Number & " - " & UserRecord.Name
            If InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "A:") = 0 And InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "D:") <> 0 Then
                If cmdAccept.Caption = "Accept" Then
                    Response = MsgBox("Do you want to make a new reservation?", vbYesNo, "HeroPOS")
                    Select Case Response
                        Case vbYes
                            NewReservation
                            Exit Sub
                        Case vbNo
                            Res_No = Val(Mid(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1, (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "-") - 1) - (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1)))
                            LoadResDetail Res_No
                            Exit Sub
                    End Select
                Else
                    Res_No = Val(Mid(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1, (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "-") - 1) - (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1)))
                    LoadResDetail Res_No
                    Exit Sub
                End If
            End If
            If InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "D:") <> 0 And InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "A:") <> 0 Then
                If cmdAccept.Caption = "Save" Or cmdAccept.Caption = "Check In" Or cmdAccept.Caption = "Check Out" Or cmdAccept.Caption = "" Then
                    LoadResDetail TillData.Res_No
                    Exit Sub
                End If
                Load frmChooseRes
                StringA = Mid(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), InStrRev(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "|"))
                stringD = Mid(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), 1, InStrRev(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "|") - 1)
                If InStr(StringA, "D") = 0 Then
                    SaveString = StringA
                    StringA = stringD
                    stringD = SaveString
                End If
                Res_TypeA = Val(Replace(Mid(StringA, 1, InStr(StringA, "<")), "|", ""))
                Res_No = Val(Mid(StringA, InStr(StringA, ":") + 1, (InStr(StringA, "-") - 1) - (InStr(StringA, ":") + 1)))
                frmChooseRes.cmdRes(0).Caption = Trim(Mid(StringA, InStrRev(StringA, "-") + 1))
                frmChooseRes.cmdRes(0).Tag = Res_No
                Select Case Res_TypeA
                    Case 0: frmChooseRes.cmdRes(0).BackColor = &HC0FFFF
                    Case 1: frmChooseRes.cmdRes(0).BackColor = &HC0FFC0
                    Case 2: frmChooseRes.cmdRes(0).BackColor = &HFFC0C0
                    Case 3: frmChooseRes.cmdRes(0).BackColor = &HC0C0FF
                End Select
                Res_TypeD = Val(Replace(Mid(stringD, 1, InStr(stringD, "<")), "|", ""))
                Res_No = Val(Mid(stringD, InStr(stringD, ":") + 1, (InStr(stringD, "-") - 1) - (InStr(stringD, ":") + 1)))
                frmChooseRes.cmdRes(1).Caption = Trim(Mid(stringD, InStrRev(stringD, "-") + 1))
                frmChooseRes.cmdRes(1).Tag = Res_No
                Select Case Res_TypeD
                    Case 0: frmChooseRes.cmdRes(1).BackColor = &HC0FFFF
                    Case 1: frmChooseRes.cmdRes(1).BackColor = &HC0FFC0
                    Case 2: frmChooseRes.cmdRes(1).BackColor = &HFFC0C0
                    Case 3: frmChooseRes.cmdRes(1).BackColor = &HC0C0FF
                End Select
                frmChooseRes.Show vbModal
                If Val(lblResNo.Tag) = 0 Then
                    Unload Me
                    Exit Sub
                End If
                LoadResDetail lblResNo.Tag
                Exit Sub
            Else
                If InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "A:") <> 0 Then
                    Res_No = Val(Mid(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1, (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "-") - 1) - (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1)))
                    LoadResDetail Res_No
                    Exit Sub
                End If
                If InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "D:") <> 0 Then
                    Res_No = Val(Mid(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1, (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "-") - 1) - (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1)))
                    LoadResDetail Res_No
                    Exit Sub
                End If
                If InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "D:") = 0 And InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "A:") = 0 Then
                    Res_No = Val(Mid(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1, (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), "-") - 1) - (InStr(.grdRes.TextMatrix(.grdRes.Row, .grdRes.Col), ":") + 1)))
                    LoadResDetail Res_No
                    Exit Sub
                End If
            End If
        Else
            NewReservation
            Exit Sub
        End If
    End With
End Sub
Private Sub NewReservation()
With frmRes
    fmType.ForeColor = &HC0C0&
    fmType.Caption = "Provisional Booking"
    frmCheckin.Caption = "Guest Check In > New Reservation"
    lblUser.Caption = UserRecord.User_Number & " - " & UserRecord.Name
    DTArrival.Value = .grdRes.TextMatrix(2, .grdRes.Col)
    DTDeparture.Value = DateAdd("D", 1, .grdRes.TextMatrix(2, .grdRes.Col))
    lblResNo.Caption = ""
    cmbRate.Enabled = True
    cmbRoomNo.Enabled = False
    txtDescription.Enabled = False
    cmdDeposit.Enabled = False
    cmbTitle.Text = "Mr."
    txtFName.Text = ""
    txtLName.Text = ""
    txtID.Text = ""
    txtVehReg.Text = ""
    txtFree.Text = "0"
    txtEmail.Text = ""
    txtTel.Text = ""
    txtMobile.Text = ""
    txtFax.Text = ""
    txtAddress.Text = ""
    txtCode.Text = ""
    txtAdults.Text = "0"
    txt5.Text = "0"
    txt0to5.Text = "0"
    txt12to16.Text = "0"
    txtContact.Text = ""
    txtContactNo.Text = ""
    txtRemarks.Text = ""
    DTArrival.Enabled = False
    DTDeparture.Enabled = True
    DTArrTime.Enabled = True
    DTDepTime.Enabled = True
    cmbTitle.Enabled = True
    txtFName.Enabled = True
    txtLName.Enabled = True
    txtID.Enabled = True
    txtVehReg.Enabled = True
    txtEmail.Enabled = True
    txtTel.Enabled = True
    txtMobile.Enabled = True
    txtFax.Enabled = True
    txtAddress.Enabled = True
    txtCode.Enabled = True
    txtAdults.Enabled = True
    cmbCountry.Enabled = True
    cmbProvince.Enabled = True
    cmbCity.Enabled = True
    txt5.Enabled = True
    txt0to5.Enabled = True
    txt12to16.Enabled = True
    txtContact.Enabled = True
    txtContactNo.Enabled = True
    txtRemarks.Enabled = True
    cmbBusiness.Enabled = True
    cmbPay.Enabled = True
    cmbBooked.Enabled = True
    Image3.BackColor = &HFFFFFF
    Image4.BackColor = &HFFFFFF
    txtNights.Enabled = True
    txtNights.Text = Val(DTDeparture.Value - DTArrival.Value)
End With
End Sub

Private Sub cmdRate_Click()
    Load frmRates
    frmRates.Tag = "Check"
    frmRates.Show vbModal
End Sub

Private Sub txt0to5_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
    Calc_Rate
End Sub

Private Sub txt0to5_GotFocus()
    txt0to5.SelStart = 0
    txt0to5.SelLength = Len(txt0to5.Text)
End Sub

Private Sub txt0to5_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtAdults.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txt0to5_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txt12to16_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
    Calc_Rate
End Sub

Private Sub txt12to16_GotFocus()
    txt12to16.SelStart = 0
    txt12to16.SelLength = Len(txt12to16.Text)
End Sub

Private Sub txt12to16_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txt5.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txt12to16_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txt5_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
    Calc_Rate
End Sub

Private Sub txt5_GotFocus()
    txt5.SelStart = 0
    txt5.SelLength = Len(txt5.Text)
End Sub

Private Sub txt5_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txt0to5.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txt5_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtAddress_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            If txtAddress.GetLineFromChar(txtAddress.SelStart) = 0 Then
                KeyCode = 0
                txtFax.SetFocus
            End If
        Case 13, 40
            If KeyCode = 40 And InStr(Mid(txtAddress.Text, txtAddress.SelStart + 1), Chr(13)) = 0 Then
                KeyCode = 0
                cmbCountry.SetFocus
            End If
            If txtAddress.GetLineFromChar(txtAddress.SelStart) = 4 Then
                KeyCode = 0
                cmbCountry.SetFocus
            End If
    End Select
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtAdults_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
    Calc_Rate
End Sub
Private Sub txtAdults_GotFocus()
    txtAdults.SelStart = 0
    txtAdults.SelLength = Len(txtAdults.Text)
End Sub
Private Sub txtAdults_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtFree.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtAdults_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            cmbCity.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtContact_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            cmbBooked.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtContactNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtContact.SetFocus
        Case 13, 40
            KeyCode = 0
            If picRates.Visible = True Then
                cmbBusiness.SetFocus
            Else
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtCredit_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                cmbPay.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
    On Error GoTo 0
End Sub
Private Sub txtCredit_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 32, 40, 41, 47, 48 To 57
        Case 39
            KeyAscii = 0
        Case 65 To 90
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtDescription_GotFocus()
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
End Sub
Private Sub txtEmail_GotFocus()
    txtEmail.SelStart = 0
    txtEmail.SelLength = Len(txtEmail.Text)
End Sub
Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtVehReg.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtFax_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
End Sub

Private Sub txtFax_GotFocus()
    txtFax.SelStart = 0
    txtFax.SelLength = Len(txtFax.Text)
End Sub
Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtMobile.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub
Private Sub txtFax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtFName_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
End Sub

Private Sub txtFName_GotFocus()
    txtFName.SelStart = 0
    txtFName.SelLength = Len(txtFName.Text)
End Sub

Private Sub txtFName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            cmbTitle.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtFName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 58
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtFree_Change()
    Calc_Rate
End Sub
Private Sub txtFree_GotFocus()
    txtFree.SelStart = 0
    txtFree.SelLength = Len(txtFree.Text)
End Sub
Private Sub txtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtCode.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub
Private Sub txtFree_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtID_GotFocus()
    txtID.SelStart = 0
    txtID.SelLength = Len(txtID.Text)
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtLName.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtLName_Change()
    Static SearchOnce
    frmCheckin.txtLName.Tag = ""
    If SearchOnce = True And txtLName.Text = "" Then SearchOnce = False
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
    If cmdAccept.Caption <> "Save" Then
        If fmType.Tag = "" Then
            If Len(txtLName.Text) > 3 And SearchOnce = False Then
                ActiveReadServer2 "Select Arrive_Date,Depart_Date, Last_Name + ' ' + First_Name + ' ' + Title as Res_Name,Res_No from Reservations where Last_Name like '" & txtLName.Text & "%' and Res_Type = 3"
                frmPrev.grdRes.Rows = 1
                While Not rs2.EOF
                    With frmPrev
                        .grdRes.Rows = .grdRes.Rows + 1
                        .grdRes.TextMatrix(.grdRes.Rows - 1, 0) = Format(rs2.Fields("Arrive_Date"), "DDD DD MMM YYYY")
                        .grdRes.TextMatrix(.grdRes.Rows - 1, 1) = Format(rs2.Fields("Depart_Date"), "DDD DD MMM YYYY")
                        .grdRes.TextMatrix(.grdRes.Rows - 1, 2) = UCase(rs2.Fields("Res_Name"))
                        .grdRes.TextMatrix(.grdRes.Rows - 1, 3) = rs2.Fields("Res_No")
                    End With
                    rs2.MoveNext
                Wend
                If rs2.RecordCount > 0 Then
                    SearchOnce = True
                    frmCheckin.txtLName.Tag = ""
                    frmPrev.Show vbModal
                    If frmCheckin.txtLName.Tag <> "" Then
                        ActiveReadServer1 "Select * from Reservations where Res_No = " & frmCheckin.txtLName.Tag
                        If rs1.RecordCount > 0 Then
                            cmbTitle.Text = rs1.Fields("Title") & ""
                            txtFName.Text = rs1.Fields("First_Name") & ""
                            txtLName.Text = rs1.Fields("Last_Name") & ""
                            txtID.Text = rs1.Fields("ID_No") & ""
                            txtVehReg.Text = rs1.Fields("Vehicle_No") & ""
                            txtEmail.Text = rs1.Fields("EMail") & ""
                            txtTel.Text = rs1.Fields("Tel_No") & ""
                            txtMobile.Text = rs1.Fields("Cell_No") & ""
                            txtFax.Text = rs1.Fields("Fax_No") & ""
                            txtAddress.Text = rs1.Fields("Address") & ""
                            txtCode.Text = rs1.Fields("Post_Code") & ""
                            cmbCountry.Text = rs1.Fields("Country") & ""
                            cmbProvince = rs1.Fields("Province") & ""
                            cmbCity = rs1.Fields("City") & ""
                            txtContact.Text = rs1.Fields("Contact_Person") & ""
                            txtContactNo.Text = rs1.Fields("Contact_no") & ""
                            cmbBusiness.Text = rs1.Fields("Source") & ""
                            txtCredit.Text = rs1.Fields("Credit_Card") & ""
                            frmCheckin.txtLName.Tag = ""
                        End If
                        rs1.Close
                    End If
                End If
                rs2.Close
            End If
        End If
    Else
        SearchOnce = False
    End If
End Sub
Private Sub txtLName_GotFocus()
    frmCheckin.txtLName.Tag = ""
    txtLName.SelStart = 0
    txtLName.SelLength = Len(txtLName.Text)
    Load frmPrev
    frmPrev.top = txtLName.top
    frmPrev.Left = txtLName.Left + txtLName.Width + 400
End Sub
Private Sub txtLName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtFName.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub
Private Sub txtLName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 58
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtMobile_GotFocus()
    txtMobile.SelStart = 0
    txtMobile.SelLength = Len(txtMobile.Text)
End Sub

Private Sub txtMobile_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtTel.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            If txtRemarks.GetLineFromChar(txtRemarks.SelStart) = 0 Then
                KeyCode = 0
                txtCredit.SetFocus
            End If
        Case 13, 40
            If KeyCode = 40 And InStr(Mid(txtRemarks.Text, txtRemarks.SelStart + 1), Chr(13)) = 0 Then
                KeyCode = 0
                SendKeys "{TAB}"
            End If
            If txtRemarks.GetLineFromChar(txtRemarks.SelStart) = 5 Then
                KeyCode = 0
                SendKeys "{TAB}"
            End If
    End Select
End Sub
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtTel_Change()
    If Validate = True Then cmdAccept.Enabled = True Else cmdAccept.Enabled = False
End Sub

Private Sub txtTel_GotFocus()
    txtTel.SelStart = 0
    txtTel.SelLength = Len(txtTel.Text)
End Sub

Private Sub txtTel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtEmail.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtVehReg_GotFocus()
    txtVehReg.SelStart = 0
    txtVehReg.SelLength = Len(txtVehReg.Text)
End Sub

Private Sub txtVehReg_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtID.SetFocus
        Case 13, 40
            KeyCode = 0
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub txtVehReg_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
