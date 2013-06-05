VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account..."
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   Icon            =   "frmAccount.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHide 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   120
      ScaleHeight     =   1125
      ScaleWidth      =   5205
      TabIndex        =   35
      Top             =   7620
      Visible         =   0   'False
      Width           =   5205
      Begin btButtonEx.ButtonEx cmdPay 
         Height          =   1065
         Left            =   0
         TabIndex        =   38
         Top             =   30
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         Appearance      =   3
         Caption         =   "Make a Payment..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdRev 
         Height          =   1065
         Left            =   1650
         TabIndex        =   39
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1879
         Appearance      =   3
         Caption         =   "Reverse GRV..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdRevPay 
         Height          =   1065
         Left            =   3330
         TabIndex        =   40
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1879
         Appearance      =   3
         Caption         =   "Reverse Payment..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTStart 
      Height          =   345
      Left            =   7800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   570
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65994755
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTStop 
      Height          =   345
      Left            =   7800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   990
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65994755
      CurrentDate     =   38862
   End
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   765
      Left            =   1530
      TabIndex        =   5
      Top             =   900
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1349
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmAccount.frx":000C
   End
   Begin VSFlex8Ctl.VSFlexGrid grdAcc 
      Height          =   4830
      Left            =   60
      TabIndex        =   2
      Top             =   2310
      Width           =   10065
      _cx             =   17754
      _cy             =   8520
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16639711
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAccount.frx":008E
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   4
         Top             =   11700
         Width           =   1005
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   3
         Top             =   5670
         Width           =   1005
      End
   End
   Begin MSComCtl2.DTPicker DTArrival 
      Height          =   345
      Left            =   1440
      TabIndex        =   23
      Top             =   7650
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65994755
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTDeparture 
      Height          =   345
      Left            =   1440
      TabIndex        =   24
      Top             =   8040
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65994755
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTDepTime 
      Height          =   345
      Left            =   3480
      TabIndex        =   25
      Top             =   8040
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   65994754
      CurrentDate     =   38862.4583333333
   End
   Begin MSComCtl2.DTPicker DTArrTime 
      Height          =   345
      Left            =   3480
      TabIndex        =   26
      Top             =   7650
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   65994754
      CurrentDate     =   38862.5833333333
   End
   Begin btButtonEx.ButtonEx cmdRevRA 
      Height          =   1035
      Left            =   4830
      TabIndex        =   42
      Top             =   7650
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1826
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Reverse Payment"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdRec 
      Height          =   1035
      Left            =   5820
      TabIndex        =   43
      Top             =   7650
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1826
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Receive Payment"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.CheckBox chkPay 
      Height          =   315
      Left            =   8370
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   1665
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2937;556"
      Value           =   "0"
      Caption         =   "Hide Paid Invoices"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comp:"
      Height          =   225
      Index           =   4
      Left            =   2550
      TabIndex        =   37
      Top             =   8460
      Width           =   795
   End
   Begin VB.Label txtFree 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3630
      TabIndex        =   36
      Top             =   8475
      Width           =   825
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7650
      TabIndex        =   34
      Top             =   8370
      Width           =   2295
   End
   Begin VB.Label lblVat 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7650
      TabIndex        =   33
      Top             =   8010
      Width           =   2295
   End
   Begin VB.Label lblSub 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7650
      TabIndex        =   32
      Top             =   7650
      Width           =   2295
   End
   Begin VB.Label lblRoom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7920
      TabIndex        =   31
      Top             =   1440
      Width           =   1845
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Departure Date:"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   27
      Top             =   8100
      Width           =   1185
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Date:"
      Height          =   255
      Index           =   2
      Left            =   390
      TabIndex        =   28
      Top             =   7710
      Width           =   975
   End
   Begin VB.Label txtNights 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      TabIndex        =   30
      Top             =   8475
      Width           =   825
   End
   Begin MSForms.Image Image8 
      Height          =   285
      Left            =   1440
      Top             =   8430
      Width           =   1065
      BorderColor     =   12632256
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1879;503"
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nights:"
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   29
      Top             =   8460
      Width           =   975
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      TabIndex        =   22
      Top             =   555
      Width           =   4275
   End
   Begin MSForms.Image Image4 
      Height          =   285
      Left            =   1440
      Top             =   510
      Width           =   4515
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "7964;503"
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name: "
      Height          =   195
      Left            =   480
      TabIndex        =   21
      Top             =   570
      Width           =   945
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Account No:"
      Height          =   195
      Left            =   480
      TabIndex        =   20
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   165
      Left            =   390
      TabIndex        =   19
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Room Number:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6210
      TabIndex        =   18
      Top             =   1410
      Width           =   1545
   End
   Begin VB.Label lblAcc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      TabIndex        =   17
      Top             =   230
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Listed From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6750
      TabIndex        =   16
      Top             =   645
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6750
      TabIndex        =   15
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Listed To:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6750
      TabIndex        =   14
      Top             =   1050
      Width           =   1005
   End
   Begin MSForms.Image Image10 
      Height          =   885
      Left            =   1440
      Top             =   840
      Width           =   4515
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "7964;1561"
   End
   Begin VB.Label lblContact 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7950
      TabIndex        =   13
      Top             =   1020
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   5805
      TabIndex        =   12
      Top             =   8430
      Width           =   1545
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vat:"
      Height          =   195
      Left            =   5805
      TabIndex        =   11
      Top             =   8055
      Width           =   1545
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5805
      TabIndex        =   10
      Top             =   7710
      Width           =   1545
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calculated Totals"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5790
      TabIndex        =   9
      Top             =   7245
      Width           =   4305
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   7230
      Width           =   3765
   End
   Begin VB.Label lblDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7890
      TabIndex        =   6
      Top             =   270
      Width           =   1995
   End
   Begin MSForms.Image Image19 
      Height          =   345
      Left            =   7800
      Top             =   1350
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "3942;609"
   End
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   7800
      Top             =   180
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3942;609"
   End
   Begin MSForms.Image Image2 
      Height          =   1755
      Left            =   6090
      Top             =   60
      Width           =   4035
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7117;3096"
   End
   Begin MSForms.Image Image13 
      Height          =   315
      Left            =   7410
      Top             =   7650
      Width           =   2655
      BorderColor     =   12632256
      Size            =   "4683;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image14 
      Height          =   315
      Left            =   7410
      Top             =   8010
      Width           =   2655
      BorderColor     =   12632256
      Size            =   "4683;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image15 
      Height          =   315
      Left            =   7410
      Top             =   8370
      Width           =   2655
      BorderColor     =   12632256
      Size            =   "4683;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   5730
      Top             =   7170
      Width           =   4425
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7805;661"
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   60
      Top             =   7170
      Width           =   5655
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "9975;661"
   End
   Begin MSForms.Image Image5 
      Height          =   1215
      Left            =   5730
      Top             =   7560
      Width           =   4425
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7805;2143"
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   1920
      Width           =   9975
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   60
      Top             =   1860
      Width           =   10065
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "17754;661"
   End
   Begin MSForms.Image Image7 
      Height          =   285
      Left            =   1440
      Top             =   180
      Width           =   2535
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "4471;503"
   End
   Begin MSForms.Image Image1 
      Height          =   1755
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   5985
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "10557;3096"
   End
   Begin MSForms.Image Image9 
      Height          =   285
      Left            =   3480
      Top             =   8430
      Width           =   1065
      BorderColor     =   12632256
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "1879;503"
   End
   Begin MSForms.Image Image3 
      Height          =   1215
      Left            =   30
      Top             =   7560
      Width           =   5655
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "9975;2143"
   End
   Begin MSForms.Image Image6 
      Height          =   8835
      Left            =   0
      Top             =   0
      Width           =   10230
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "18045;15584"
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkPay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEvents
    Select Case chkPay.Value
        Case False
            For i = 1 To grdAcc.Rows - 1
                If grdAcc.FindRow(grdAcc.TextMatrix(i, 5), 0, 6) <> -1 Then
                    MyRow = grdAcc.FindRow(grdAcc.TextMatrix(i, 5), 0, 6)
                    If grdAcc.ValueMatrix(i, 3) = grdAcc.ValueMatrix(MyRow, 2) Then
                        grdAcc.RowHidden(i) = True
                        grdAcc.RowHidden(MyRow) = True
                    End If
                End If
            Next i
            Exit Sub
        Case True
            For i = 1 To grdAcc.Rows - 1
                grdAcc.RowHidden(i) = False
            Next i
            Exit Sub
    End Select
End Sub
Private Sub cmdPay_Click()
    Load frmPayment
    frmPayment.Tag = "Account"
    frmPayment.Show vbModal
    frmPayment.Tag = ""
End Sub

Private Sub cmdRec_Click()
    Load frmRecPayment
    frmRecPayment.Tag = "Debtor"
    frmRecPayment.Show vbModal
End Sub

Private Sub cmdRevPay_Click()
    ActiveReadServer "Select Balance from Supplier_Accounts where Account_No = '" & lblAcc.Caption & "' order by Line_No"
    If rs.RecordCount > 0 Then
        rs.MoveLast
        Balance = rs.Fields("Balance")
    End If
    rs.Close
    ActiveUpdateServer "Delete from Supplier_Accounts where Payment_No = " & Val(Trim(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), InStr(grdAcc.TextMatrix(grdAcc.Row, 1), "-") + 1, 7)))
    DoEvents
    ActiveUpdateServer "Update Suppliers set Balance= Balance + " & grdAcc.TextMatrix(grdAcc.Row, 2) & " where Supplier_No='" & lblAcc.Caption & "'"
    Form_Activate
    MsgBox "Payment Removal Complete", vbInformation, "HeroPOS"
End Sub
Private Sub cmdRevRA_Click()
    If cmdRevRA.Caption = "Transfer to Room" Then
        frmRoomTransfer.Show vbModal
        Select Case frmRoomTransfer.Tag
            Case ""
            Case Else
                ActiveReadServer "Select Res_No from Reservations where Room_No = " & frmRoomTransfer.Tag & " and Res_Type = 2"
                If rs.RecordCount > 0 Then
                    Res_No = rs.Fields("Res_No")
                End If
                rs.Close
                ActiveUpdateServer "Update Room_Accounts set Res_No = '" & Res_No & "' where Res_No = '" & Val(lblAcc.Caption) & "' and Invoice_No = " & Val(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), InStrRev(grdAcc.TextMatrix(grdAcc.Row, 1), "-") + 1))
                DoEvents
                Screen.MousePointer = 0
                MsgBox "Transaction Transfer Completed.", vbInformation, "HeroPOS"
        End Select
        Unload frmRoomTransfer
        Form_Activate
        Exit Sub
    End If
    If cmdRevRA.Caption = "Transfer to Debtor" Then
        frmDebTrans.Show vbModal
        Select Case frmDebTrans.Tag
            Case ""
            Case Else
                Screen.MousePointer = 11
                ActiveUpdateServer "Update Debtor_Accounts set Account_No = '" & frmDebTrans.Tag & "' where Account_No = '" & lblAcc.Caption & "' and Invoice_No = " & Val(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), InStrRev(grdAcc.TextMatrix(grdAcc.Row, 1), "-") + 1))
                DoEvents
                ActiveUpdateServer "Update Sales_Journal set Account_No = '" & frmDebTrans.Tag & "' where Invoice_no = " & Val(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), InStrRev(grdAcc.TextMatrix(grdAcc.Row, 1), "-") + 1))
                
                DoEvents
                
                ActiveReadServer2 "Select * from Debtor_Accounts where Account_No = '" & frmDebTrans.Tag & "' order by Date_Time"
                Balance = 0
                While Not rs2.EOF
                    Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
                    ActiveUpdateServer "Update Debtor_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
                    rs2.MoveNext
                Wend
                rs2.Close
                ActiveUpdateServer "Update Debtors set Balance = " & Balance & " Where Debtor_no = '" & frmDebTrans.Tag & "'"
                
                DoEvents
                
                ActiveReadServer2 "Select * from Debtor_Accounts where Account_No = '" & lblAcc.Caption & "' order by Date_Time"
                Balance = 0
                While Not rs2.EOF
                    Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
                    ActiveUpdateServer "Update Debtor_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
                    rs2.MoveNext
                Wend
                rs2.Close
                
                ActiveUpdateServer "Update Debtors set Balance = " & Balance & " Where Debtor_no = '" & lblAcc.Caption & "'"
                
                Screen.MousePointer = 0
                MsgBox "Transaction Transfer Completed.", vbInformation, "HeroPOS"
        End Select
        Unload frmDebTrans
        Exit Sub
    End If
    If cmdRevRA.Caption = "Reverse Deposit" Then
        Screen.MousePointer = 11
        ActiveUpdateServer "Delete from Room_Accounts where Res_No = '" & Val(lblAcc.Caption) & "' and Transaction_Type = 'Deposit' and Invoice_No = " & Val(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), InStrRev(grdAcc.TextMatrix(grdAcc.Row, 1), "-") + 1))
        MsgBox "Deposit Reversal Completed.", vbInformation, "HeroPOS"
        Form_Activate
        Screen.MousePointer = 0
        Exit Sub
    End If
    Select Case frmAccount.Tag
        Case "Debtor"
            ActiveReadServer "Select Balance from Debtor_Accounts where Account_No = '" & lblAcc.Caption & "' order by Line_No"
            If rs.RecordCount > 0 Then
                rs.MoveLast
                Balance = rs.Fields("Balance")
            End If
            rs.Close
            ActiveUpdateServer "Delete from Debtor_Accounts where Payment_No = " & Val(Trim(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), 21, 7)))
            DoEvents
            ActiveReadServer2 "Select * from Debtor_Accounts where Account_No = '" & lblAcc.Caption & "' order by Date_Time"
            Balance = 0
            While Not rs2.EOF
                Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
                ActiveUpdateServer "Update Debtor_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
                rs2.MoveNext
            Wend
            rs2.Close
            Form_Activate
            MsgBox "Payment Removal Complete", vbInformation, "HeroPOS"
        Case Else
            ActiveUpdateServer "Delete from Room_Accounts where Transaction_Type ='Receipt' and Res_no= " & TillData.Res_No & " and Invoice_No = " & Val(Trim(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), InStr(grdAcc.TextMatrix(grdAcc.Row, 1), "-") + 1, 7)))
            DoEvents
            Form_Activate
    End Select
End Sub

Private Sub DTStart_CloseUp()
    Form_Activate
End Sub
Private Sub DTStop_CloseUp()
    Form_Activate
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 11
    DoEvents
    If grdAcc.Tag = "1" Then
        grdAcc.Tag = ""
        Screen.MousePointer = 0
        Exit Sub
    End If
    lblSub.Caption = "0.00"
    If frmAccount.Tag = "Debtor" Then
        chkPay.Visible = True
        cmdPay.Visible = False
        lblRoom.Visible = False
        picHide.Visible = False
        ActiveReadServer "Select * from Debtors where Debtor_No = '" & TillData.Account_No & "'"
        If rs.RecordCount > 0 Then
            lblAcc.Caption = TillData.Account_No
            lblName.Caption = rs.Fields("Debtor_Name")
            txtAddress.Text = rs.Fields("Address")
            lblDate.Caption = Format(Date, "ddd dd MMM yyyy")
            lblRoom.Caption = ""
            txtNights.Caption = ""
            lblInfo.Caption = " Credit Limit: " & Format(rs.Fields("Credit_Limit"), "0.00") & " - "
            Select Case rs.Fields("Debt_Type")
                Case 0: lblInfo.Caption = lblInfo.Caption & " (Debtor)"
                Case 1: lblInfo.Caption = lblInfo.Caption & " (Staff Account)"
                Case 2: lblInfo.Caption = lblInfo.Caption & " (Management Account)"
            End Select
        End If
        rs.Close
    End If
    If frmAccount.Tag = "Supplier" Then
        chkPay.Visible = True
        cmdPay.Visible = True
        picHide.Visible = True
        ActiveReadServer "Select * from Suppliers where Supplier_No = '" & TillData.Account_No & "'"
        If rs.RecordCount > 0 Then
            lblAcc.Caption = TillData.Account_No
            lblName.Caption = rs.Fields("Supplier_Name")
            txtAddress.Text = rs.Fields("Address")
            lblDate.Caption = Format(Date, "ddd dd MMM yyyy")
            lblRoom.Visible = False
            txtNights.Caption = ""
            lblInfo.Caption = " Credit Limit: " & Format(rs.Fields("Credit_Limit"), "0.00") & " - "
            lblInfo.Caption = lblInfo.Caption & " (Supplier)"
        End If
        rs.Close
    End If
    If frmAccount.Tag = "Room" Then
        cmdPay.Visible = False
        lblRoom.Visible = True
        picHide.Visible = False
        ActiveReadServer "Select * from Reservations where Res_No = " & TillData.Res_No
        If rs.RecordCount > 0 Then
            lblAcc.Caption = Format(TillData.Res_No, "000000")
            lblName.Caption = rs.Fields("Title") & " " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
            txtAddress.Text = rs.Fields("Address")
            lblDate.Caption = Format(Date, "ddd dd MMM yyyy")
            DTStart.Value = rs.Fields("Arrive_Date")
            DTStop.Value = rs.Fields("Depart_Date")
            DTArrival.Value = rs.Fields("Arrive_Date")
            DTDeparture.Value = rs.Fields("Depart_Date")
            lblRoom.Caption = rs.Fields("Room_No")
            txtFree.Caption = Val(rs.Fields("Free_Nights") & "")
            txtNights.Caption = Val(DTDeparture.Value - DTArrival.Value)
            lblInfo.Caption = " Accomodation: " & txtNights.Caption & " Nights @ " & Format(rs.Fields("Rate"), "0.00") & " = " & Format(rs.Fields("Rate") * (Val(txtNights.Caption) - Val(txtFree.Caption)), "0.00")
            Select Case rs.Fields("Res_Type")
                Case 0: lblInfo.Caption = lblInfo.Caption & " (Provisional)"
                Case 1: lblInfo.Caption = lblInfo.Caption & " (Confirmed)"
                Case 2: lblInfo.Caption = lblInfo.Caption & " (Checked In)"
                Case 3: lblInfo.Caption = lblInfo.Caption & " (Checked Out)"
            End Select
        End If
        rs.Close
    End If
    Select Case frmAccount.Tag
        Case "Supplier"
            ActiveReadServer "Select 1 as Line_No,'" & DTStart.Value & "' as Date_Time,'' as Invoice_No,'' as Payment_No,'Opening Balance' as Transaction_Type" & _
            " ,Account_No,Sum(Debit)as Debit,Sum(Credit)as Credit,Sum(Debit)-Sum(Credit) as Balance," & _
            " 0 as User_No,'' as Tender_Type,'' as Ref_No" & _
            " from Supplier_Accounts where Account_No = '" & lblAcc.Caption & "'" & _
            " and (Date_Time < '" & DTStart.Value & "' )" & _
            " group by Account_No " & _
            " Union" & _
            " Select * from Supplier_Accounts where Account_No = '" & lblAcc.Caption & "'" & _
            " and (Date_Time > '" & DTStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & DTStop.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "') order by Date_Time"
            lblType.Caption = "Account No:"
        Case "Debtor"
            If Zero_Print = 0 Then
                ActiveReadServer "Select 1 as Line_No,'" & DTStart.Value & "' as Date_Time,0 as Invoice_No,'Opening Balance' as Transaction_Type" & _
                " ,Account_No,Sum(Debit)as Debit,Sum(Credit)as Credit,Sum(Debit)-Sum(Credit) as Balance," & _
                " 0 as User_No,'' as Tender_Type,'' as Ref_No,'' as Payment_No" & _
                " from Debtor_Accounts where Account_No = '" & TillData.Account_No & "'" & _
                " and (Date_Time < '" & DTStart.Value & "' )" & _
                " group by Account_No" & _
                " Union" & _
                " Select * from Debtor_Accounts where Account_No = '" & TillData.Account_No & "'" & _
                " and (Date_Time > '" & DateAdd("d", -1, DTStart.Value) & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and Date_Time<'" & DTStop.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "') order by Date_Time"
            Else
                ActiveReadServer "Select 1 as Line_No,'" & DTStart.Value & "' as Date_Time,0 as Invoice_No,'Opening Balance' as Transaction_Type" & _
                " ,Account_No,Sum(Debit)as Debit,Sum(Credit)as Credit,Sum(Debit)-Sum(Credit) as Balance," & _
                " 0 as User_No,'' as Tender_Type,'' as Ref_No,'' as Payment_No" & _
                " from Debtor_Accounts where Account_No = '" & TillData.Account_No & "'" & _
                " and (Date_Time < '" & DTStart.Value & "' )" & _
                " group by Account_No" & _
                " Union" & _
                " Select * from Debtor_Accounts where Account_No = '" & TillData.Account_No & "' and Debit-Credit <> 0" & _
                " and (Date_Time > '" & DateAdd("d", -1, DTStart.Value) & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and Date_Time<'" & DTStop.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "') order by Date_Time"
            End If
            lblType.Caption = "Account No:"
        Case "Room"
            ActiveReadServer "Select * from Room_Accounts where Res_No = '" & TillData.Res_No & "' order by Date_Time"
            lblType.Caption = "Reservation No:"
    End Select
    Balance = 0
    Totvat = 0
    grdAcc.Rows = 1
    While Not rs.EOF
        grdAcc.Rows = grdAcc.Rows + 1
        grdAcc.Row = grdAcc.Rows - 1
        If rs.Fields("Transaction_Type") & "" = "Opening Balance" Then
            grdAcc.TextMatrix(grdAcc.Row, 5) = ""
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            grdAcc.TextMatrix(grdAcc.Row, 1) = "Balance Brought Forward"
            grdAcc.TextMatrix(grdAcc.Row, 2) = "0.00"
            grdAcc.TextMatrix(grdAcc.Row, 3) = "0.00"
            Balance = Balance + Format(rs.Fields("Balance") & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(rs.Fields("Balance") & "", "0.00")
        End If
        If rs.Fields("Transaction_Type") & "" = "Accomodation" Then
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            grdAcc.TextMatrix(grdAcc.Row, 1) = "Accomodation Sales Invoice - " & Format(rs.Fields("Invoice_No"), "000000")
            grdAcc.TextMatrix(grdAcc.Row, 2) = Format(rs.Fields("Debit") & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 3) = "0.00"
            Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
            Totvat = Totvat + rs.Fields("Debit") - (rs.Fields("Debit") / 1.14)
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(Balance & "", "0.00")
        End If
        If rs.Fields("Transaction_Type") & "" = "Supplier Invoice" Then
            ActiveReadServer1 "Select Invoice_No from Purchase_Journal where Invoice_No is not Null and GRV_No = " & Val(rs.Fields("Invoice_No"))
            If rs1.RecordCount > 0 Then
                Invoice_No = rs1.Fields("Invoice_No") & ""
                grdAcc.TextMatrix(grdAcc.Row, 5) = Invoice_No
            End If
            rs1.Close
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            grdAcc.TextMatrix(grdAcc.Row, 1) = "Supplier Invoice - " & Invoice_No & " > GRV No: " & Format(rs.Fields("Invoice_No"), "000000")
            grdAcc.TextMatrix(grdAcc.Row, 2) = "0.00"
            grdAcc.TextMatrix(grdAcc.Row, 3) = Format(rs.Fields("Credit") & "", "0.00")
            Balance = Balance - Format(rs.Fields("Credit") & "", "0.00")
            Totvat = Totvat + rs.Fields("Credit") - (rs.Fields("Credit") / 1.14)
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(Balance & "", "0.00")
        End If
        If rs.Fields("Transaction_Type") & "" = "Payment" Then
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            grdAcc.TextMatrix(grdAcc.Row, 1) = "Supplier Payment - " & Format(rs.Fields("Payment_No"), "000000") & " > " & rs.Fields("Tender_Type") & " (Ref: " & rs.Fields("Ref_No") & ") for Invoice: " & rs.Fields("Invoice_No")
            grdAcc.TextMatrix(grdAcc.Row, 6) = rs.Fields("Invoice_No") & ""
            grdAcc.TextMatrix(grdAcc.Row, 2) = Format(rs.Fields("Debit") & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 3) = "0.00"
            Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(Balance & "", "0.00")
        End If
        If rs.Fields("Transaction_Type") & "" = "Invoice" Then
            grdAcc.TextMatrix(grdAcc.Row, 5) = rs.Fields("Invoice_No")
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            Payment_No = 0
            ActiveReadServer1 "Select Payment_No from Debtor_Accounts where Payment_No is not Null and Invoice_No = " & Val(rs.Fields("Invoice_No"))
            If rs1.RecordCount > 0 Then
                Payment_No = rs1.Fields("Payment_No") & ""
            End If
            rs1.Close
            If Payment_No = 0 Then
                    grdAcc.TextMatrix(grdAcc.Row, 1) = "Sales Invoice - " & String(7 - Len(rs.Fields("Invoice_No")), "0") & rs.Fields("Invoice_No")
                Else
                    grdAcc.TextMatrix(grdAcc.Row, 1) = "Sales Invoice - " & String(7 - Len(rs.Fields("Invoice_No")), "0") & rs.Fields("Invoice_No") & " > Payment No: " & Format(Payment_No, "000000")
            End If
            grdAcc.TextMatrix(grdAcc.Row, 2) = Format(rs.Fields("Debit") & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 3) = "0.00"
            Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
            Totvat = Totvat + rs.Fields("Debit") - (rs.Fields("Debit") / 1.14)
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(Balance & "", "0.00")
        End If
        If rs.Fields("Transaction_Type") & "" = "Receipt" Then
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            If frmAccount.Tag = "Room" Then
                grdAcc.TextMatrix(grdAcc.Row, 1) = "Receive on Account - " & Format(rs.Fields("Invoice_No"), "000000")
            Else
                grdAcc.TextMatrix(grdAcc.Row, 1) = "Receive on Account - " & Format(rs.Fields("Payment_No"), "000000") & " > " & rs.Fields("Tender_Type") & " (Ref: " & rs.Fields("Ref_No") & ") for Invoice: " & rs.Fields("Invoice_No")
            End If
            grdAcc.TextMatrix(grdAcc.Row, 3) = Format(rs.Fields("Credit") & "", "0.00")
            Balance = Balance + Format(rs.Fields("Credit") * -1 & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 2) = "0.00"
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(Balance & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 6) = rs.Fields("Invoice_No") & ""
        End If
        If rs.Fields("Transaction_Type") & "" = "Deposit" Then
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            grdAcc.TextMatrix(grdAcc.Row, 1) = "Deposit Received- " & Format(rs.Fields("Invoice_No"), "000000")
            grdAcc.TextMatrix(grdAcc.Row, 3) = Format(rs.Fields("Credit") * -1 & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 2) = "0.00"
            Balance = Balance + Format(rs.Fields("Credit") * -1 & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(Balance & "", "0.00")
        End If
        If rs.Fields("Transaction_Type") & "" = "Journal" Then
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            If rs.Fields("Debit") <> 0 Then
                grdAcc.TextMatrix(grdAcc.Row, 1) = "Journal for " & rs.Fields("Ref_No") & " - " & Format(rs.Fields("Invoice_No"), "000000")
                grdAcc.TextMatrix(grdAcc.Row, 2) = Format(rs.Fields("Debit"), "0.00")
                grdAcc.TextMatrix(grdAcc.Row, 3) = "0.00"
                Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
            Else
                grdAcc.TextMatrix(grdAcc.Row, 1) = "Journal for " & rs.Fields("Ref_No") & " - " & Format(rs.Fields("Invoice_No"), "000000")
                grdAcc.TextMatrix(grdAcc.Row, 3) = Format(rs.Fields("Credit"), "0.00")
                grdAcc.TextMatrix(grdAcc.Row, 2) = "0.00"
                Balance = Balance - Format(rs.Fields("Credit") & "", "0.00")
            End If
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(Balance & "", "0.00")
        End If
        If Left(rs.Fields("Transaction_Type"), 9) & "" = "Telephone" Then
            grdAcc.TextMatrix(grdAcc.Row, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy")
            grdAcc.TextMatrix(grdAcc.Row, 1) = rs.Fields("Transaction_Type")
            grdAcc.TextMatrix(grdAcc.Row, 2) = Format(rs.Fields("Debit") & "", "0.00")
            grdAcc.TextMatrix(grdAcc.Row, 3) = "0.00"
            Balance = Balance + Format(rs.Fields("Debit") & "", "0.00")
            Totvat = Totvat + rs.Fields("Debit") - (rs.Fields("Debit") / 1.14)
            grdAcc.TextMatrix(grdAcc.Row, 4) = Format(Balance & "", "0.00")
        End If
        rs.MoveNext
    Wend
    grdAcc.ShowCell grdAcc.Row, 0
    rs.Close
    lblTotal.Caption = Format(Balance, "0.00")
    If Val(lblTotal.Caption) <> 0 Then
        If Balance > 0 Then
        lblSub.Caption = Format(Balance - Totvat, "0.00")
        End If
        If Balance < 0 Then
        lblSub.Caption = Format(Balance + Totvat, "0.00")
        End If
    End If
    If Val(lblSub.Caption) = 0 Then lblVat.Caption = 0
    lblVat.Caption = Format(Totvat, "0.00")
    
    If grdAcc.Rows > 1 Then grdAcc.Row = 1
    DoEvents
    
    Screen.MousePointer = 0
    If Val(lblTotal.Caption) = 0 Then lblVat.Caption = 0
    lblVat.Caption = Format(lblVat, "0.00")
End Sub
Private Sub Form_Load()
    grdAcc.Rows = 1
    grdAcc.Cols = 7
    grdAcc.TextMatrix(0, 0) = "Dated"
    grdAcc.TextMatrix(0, 1) = "Transaction Details"
    grdAcc.TextMatrix(0, 2) = "Credit"
    grdAcc.TextMatrix(0, 3) = "Debit"
    grdAcc.TextMatrix(0, 4) = "Balance"
    grdAcc.ColAlignment(0) = flexAlignLeftCenter
    grdAcc.ColAlignment(1) = flexAlignLeftCenter
    grdAcc.ColAlignment(2) = flexAlignRightCenter
    grdAcc.ColAlignment(3) = flexAlignRightCenter
    grdAcc.ColAlignment(4) = flexAlignRightCenter
    grdAcc.ColWidth(0) = grdAcc.Width * 0.1
    grdAcc.ColWidth(1) = grdAcc.Width * 0.55
    grdAcc.ColWidth(2) = grdAcc.Width * 0.11
    grdAcc.ColWidth(3) = grdAcc.Width * 0.11
    grdAcc.ColWidth(4) = grdAcc.Width * 0.11
    grdAcc.ColHidden(5) = True
    grdAcc.ColHidden(6) = True
    YYYY = Format(Date, "YYYY")
    MM = Format(Date, "MM")
    DTStart.Value = YYYY & "-" & MM & "-01"
    DTStop.Value = Date
End Sub
Private Sub grdAcc_DblClick()
    If frmAccount.Tag = "Supplier" Then
        MsgBox "Reprint or View this Transaction from the Purchase Journal", vbInformation, "HeroPOS"
        Exit Sub
    Else
        grdAcc.Tag = "1"
        Load frmTransView
        frmTransView.Tag = Val(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), InStr(grdAcc.TextMatrix(grdAcc.Row, 1), "-") + 1))
        frmTransView.Show vbModal
    End If
End Sub
Private Sub grdAcc_RowColChange()
    On Error Resume Next
    cmdRevRA.Caption = "Reverse Payment"
    cmdPay.Enabled = False
    cmdRev.Enabled = False
    cmdRevPay.Enabled = False
    cmdRevRA.Enabled = False
    For i = 1 To grdAcc.Rows - 1
        grdAcc.Cell(flexcpFontBold, i, 0, i, 4) = False
        grdAcc.Cell(flexcpForeColor, i, 0, i, 4) = vbBlack
    Next i
    If Trim(grdAcc.TextMatrix(grdAcc.Row, 1)) <> "" And grdAcc.Rows > 1 Then
        Select Case Trim(Mid(grdAcc.TextMatrix(grdAcc.Row, 1), 1, InStr(grdAcc.TextMatrix(grdAcc.Row, 1), "-") - 1))
            Case "Supplier Invoice", "Journal"
                cmdRec.Enabled = False
                If grdAcc.FindRow(grdAcc.TextMatrix(grdAcc.Row, 5), 0, 6) <> -1 Then
                    MyRow = grdAcc.FindRow(grdAcc.TextMatrix(grdAcc.Row, 5), 0, 6)
                    grdAcc.Cell(flexcpFontBold, grdAcc.Row, 0, grdAcc.Row, 4) = True
                    grdAcc.Cell(flexcpFontBold, MyRow, 0, MyRow, 4) = True
                    grdAcc.Cell(flexcpForeColor, grdAcc.Row, 0, grdAcc.Row, 4) = vbBlue
                    grdAcc.Cell(flexcpForeColor, MyRow, 0, MyRow, 4) = vbBlue
                Else
                    cmdPay.Enabled = True
                    cmdRev.Enabled = True
                End If
            Case "Supplier Payment"
                cmdRec.Enabled = False
                If grdAcc.FindRow(grdAcc.TextMatrix(grdAcc.Row, 5), 0, 6) <> -1 Then
                    MyRow = grdAcc.FindRow(grdAcc.TextMatrix(grdAcc.Row, 5), 0, 6)
                    grdAcc.Cell(flexcpFontBold, grdAcc.Row, 0, grdAcc.Row, 4) = True
                    grdAcc.Cell(flexcpFontBold, MyRow, 0, MyRow, 4) = True
                    grdAcc.Cell(flexcpForeColor, grdAcc.Row, 0, grdAcc.Row, 4) = vbBlue
                    grdAcc.Cell(flexcpForeColor, MyRow, 0, MyRow, 4) = vbBlue
                End If
                cmdRevPay.Enabled = True
            Case "Receive on Account"
                If grdAcc.FindRow(grdAcc.TextMatrix(grdAcc.Row, 6), 0, 5) <> -1 Then
                    MyRow = grdAcc.FindRow(grdAcc.TextMatrix(grdAcc.Row, 6), 0, 5)
                    grdAcc.Cell(flexcpFontBold, grdAcc.Row, 0, grdAcc.Row, 4) = True
                    grdAcc.Cell(flexcpFontBold, MyRow, 0, MyRow, 4) = True
                    grdAcc.Cell(flexcpForeColor, grdAcc.Row, 0, grdAcc.Row, 4) = vbBlue
                    grdAcc.Cell(flexcpForeColor, MyRow, 0, MyRow, 4) = vbBlue
                End If
                cmdRec.Enabled = False
                cmdRevRA.Enabled = True
            Case "Sales Invoice", "Accomodation Sales Invoice"
                If grdAcc.FindRow(grdAcc.TextMatrix(grdAcc.Row, 5), 0, 6) <> -1 Then
                    MyRow = grdAcc.FindRow(grdAcc.TextMatrix(grdAcc.Row, 5), 0, 6)
                    grdAcc.Cell(flexcpFontBold, grdAcc.Row, 0, grdAcc.Row, 4) = True
                    grdAcc.Cell(flexcpFontBold, MyRow, 0, MyRow, 4) = True
                    grdAcc.Cell(flexcpForeColor, grdAcc.Row, 0, grdAcc.Row, 4) = vbBlue
                    grdAcc.Cell(flexcpForeColor, MyRow, 0, MyRow, 4) = vbBlue
                End If
                If InStr(grdAcc.TextMatrix(grdAcc.Row, 1), "Payment") = 0 Then
                    cmdRec.Enabled = True
                    If Val(lblRoom.Caption) = 0 Then
                        cmdRevRA.Caption = "Transfer to Debtor"
                    Else
                        cmdRevRA.Caption = "Transfer to Room"
                    End If
                    cmdRevRA.Enabled = True
                Else
                    cmdRevRA.Enabled = True
                End If
            Case "Deposit Received"
                cmdRec.Enabled = False
                cmdRevRA.Caption = "Reverse Deposit"
                cmdRevRA.Enabled = True
        End Select
        If Left(grdAcc.TextMatrix(grdAcc.Row, 1), 7) = "Journal" And frmAccount.Tag = "Supplier" And grdAcc.ValueMatrix(grdAcc.Row, 3) <> 0 Then
            cmdPay.Enabled = True
        End If
    End If
    On Error GoTo 0
End Sub



