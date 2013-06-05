VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmGuests 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer timerwindow 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11520
      Top             =   1050
   End
   Begin VSFlex8Ctl.VSFlexGrid grdSupp 
      Height          =   3405
      Left            =   0
      TabIndex        =   17
      Top             =   5430
      Width           =   13740
      _cx             =   24236
      _cy             =   6006
      Appearance      =   2
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColorSel    =   15329975
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGuests.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
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
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
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
   End
   Begin VB.TextBox txtSuppName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2100
      TabIndex        =   1
      Top             =   1140
      Width           =   3585
   End
   Begin VB.TextBox txtGL_Code 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8070
      TabIndex        =   10
      Top             =   1830
      Width           =   2895
   End
   Begin VB.TextBox txtWeb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8070
      TabIndex        =   9
      Top             =   1500
      Width           =   2895
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8070
      TabIndex        =   8
      Top             =   1140
      Width           =   2895
   End
   Begin VB.TextBox txtVAT 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8070
      TabIndex        =   36
      Text            =   "<Not Set>"
      Top             =   780
      Width           =   2895
   End
   Begin VB.TextBox txtFax 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2100
      TabIndex        =   6
      Top             =   2940
      Width           =   2745
   End
   Begin VB.TextBox txtCell 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2100
      TabIndex        =   5
      Top             =   2580
      Width           =   2745
   End
   Begin VB.TextBox txtBussTell 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2100
      TabIndex        =   4
      Top             =   2220
      Width           =   2745
   End
   Begin VB.TextBox txtCredit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2100
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   1860
      Width           =   1455
   End
   Begin VB.TextBox txtContact 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2100
      TabIndex        =   2
      Top             =   1500
      Width           =   3585
   End
   Begin VB.TextBox txtSuppNo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2100
      TabIndex        =   0
      Top             =   780
      Width           =   1785
   End
   Begin VB.Frame Terms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Terms"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   7980
      TabIndex        =   21
      Top             =   2220
      Width           =   3015
      Begin VB.OptionButton optTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cash Debtor"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   390
         TabIndex        =   39
         Top             =   1470
         Width           =   1305
      End
      Begin VB.OptionButton optTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "120 Days +"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1710
         TabIndex        =   16
         Top             =   1110
         Width           =   1155
      End
      Begin VB.OptionButton optTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "120 Days"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1710
         TabIndex        =   15
         Top             =   720
         Width           =   1005
      End
      Begin VB.OptionButton optTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "90 Days"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1710
         TabIndex        =   14
         Top             =   330
         Width           =   1005
      End
      Begin VB.OptionButton optTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "60 Days "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   13
         Top             =   1110
         Width           =   1005
      End
      Begin VB.OptionButton optTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "30 Days "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Current"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   11
         Top             =   330
         Width           =   1005
      End
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   14865
      TabIndex        =   18
      Top             =   4650
      Width           =   14865
      Begin btButtonEx.ButtonEx cmdAccount 
         Height          =   465
         Left            =   11550
         TabIndex        =   37
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "View Account"
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
      Begin btButtonEx.ButtonEx cmdUp 
         Height          =   465
         Left            =   13110
         TabIndex        =   19
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "5"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdStatement 
         Height          =   465
         Left            =   7950
         TabIndex        =   38
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "Print Statement"
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
      Begin MSComCtl2.DTPicker DTStart 
         Height          =   285
         Left            =   5670
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   120
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   65798147
         CurrentDate     =   38862
      End
      Begin MSComCtl2.DTPicker DTStop 
         Height          =   285
         Left            =   5670
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   390
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   65798147
         CurrentDate     =   38862
      End
      Begin btButtonEx.ButtonEx cmdAll 
         Height          =   465
         Left            =   9300
         TabIndex        =   45
         Top             =   180
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "All..."
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
      Begin btButtonEx.ButtonEx cmdEmail 
         Height          =   465
         Left            =   10110
         TabIndex        =   46
         Top             =   180
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "Email Statements"
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
      Begin btButtonEx.ButtonEx cmdPrintpricelist 
         Height          =   465
         Left            =   7950
         TabIndex        =   48
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "View Pricelist"
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Statement to:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4140
         TabIndex        =   43
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Statement from:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4140
         TabIndex        =   42
         Top             =   165
         Width           =   1485
      End
      Begin MSForms.Image Image6 
         Height          =   90
         Left            =   60
         Top             =   570
         Width           =   3015
         BackColor       =   16761024
         Size            =   "5318;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   0
         Left            =   3090
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   1
         Left            =   3420
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   2
         Left            =   3750
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin VB.Label Label5 
         Caption         =   "Account Details..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
      Begin MSForms.Image picSeperate1 
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   13755
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "24262;185"
      End
      Begin MSForms.Image Image7 
         Height          =   675
         Left            =   0
         Top             =   90
         Width           =   13740
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "24236;1191"
      End
   End
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   1125
      Left            =   2160
      TabIndex        =   7
      Top             =   3360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1984
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmGuests.frx":0078
   End
   Begin btButtonEx.ButtonEx cmdDiscount 
      Height          =   315
      Left            =   6180
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Discount Structure..."
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
      CaptionAlignHorz=   0
   End
   Begin btButtonEx.ButtonEx cmdChange 
      Height          =   315
      Left            =   3990
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   720
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Change Debtor No..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignHorz=   0
   End
   Begin btButtonEx.ButtonEx cmdPricelist 
      Height          =   315
      Left            =   7980
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "This Debtor's Pricelist..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignHorz=   0
   End
   Begin btButtonEx.ButtonEx cmdPrintselected 
      Height          =   315
      Left            =   9870
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4260
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Print only selected fields..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignHorz=   0
   End
   Begin MSForms.Image Image16 
      Height          =   315
      Left            =   7980
      Top             =   1800
      Width           =   3015
      BackColor       =   16777215
      Size            =   "5318;556"
   End
   Begin MSForms.Image Image15 
      Height          =   315
      Left            =   7980
      Top             =   1440
      Width           =   3015
      BackColor       =   16777215
      Size            =   "5318;556"
   End
   Begin MSForms.Image Image14 
      Height          =   315
      Left            =   7980
      Top             =   1080
      Width           =   3015
      BackColor       =   16777215
      Size            =   "5318;556"
   End
   Begin MSForms.Image Image13 
      Height          =   315
      Left            =   7980
      Top             =   720
      Width           =   3015
      BackColor       =   16777215
      Size            =   "5318;556"
   End
   Begin MSForms.Image Image12 
      Height          =   315
      Left            =   2010
      Top             =   2880
      Width           =   2865
      BackColor       =   16777215
      Size            =   "5054;556"
   End
   Begin MSForms.Image Image11 
      Height          =   315
      Left            =   2010
      Top             =   2520
      Width           =   2865
      BackColor       =   16777215
      Size            =   "5054;556"
   End
   Begin MSForms.Image Image10 
      Height          =   315
      Left            =   2010
      Top             =   2160
      Width           =   2865
      BackColor       =   16777215
      Size            =   "5054;556"
   End
   Begin MSForms.Image Image9 
      Height          =   315
      Left            =   2010
      Top             =   1800
      Width           =   1575
      BackColor       =   16777215
      Size            =   "2778;556"
   End
   Begin MSForms.Image Image8 
      Height          =   315
      Left            =   2010
      Top             =   1440
      Width           =   3705
      BackColor       =   16777215
      Size            =   "6535;556"
   End
   Begin MSForms.Image Image2 
      Height          =   315
      Left            =   2010
      Top             =   1080
      Width           =   3705
      BackColor       =   16777215
      Size            =   "6535;556"
   End
   Begin MSForms.Image Image1 
      Height          =   315
      Left            =   2010
      Top             =   720
      Width           =   1905
      BackColor       =   16777215
      Size            =   "3360;556"
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit: "
      Height          =   165
      Index           =   0
      Left            =   750
      TabIndex        =   34
      Top             =   1890
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Web Page: "
      Height          =   225
      Index           =   0
      Left            =   6780
      TabIndex        =   33
      Top             =   1530
      Width           =   1155
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: "
      Height          =   165
      Left            =   6780
      TabIndex        =   32
      Top             =   1170
      Width           =   1155
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address: "
      Height          =   165
      Left            =   750
      TabIndex        =   31
      Top             =   3330
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number: "
      Height          =   165
      Left            =   750
      TabIndex        =   30
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile: "
      Height          =   165
      Left            =   750
      TabIndex        =   29
      Top             =   2610
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business Tel: "
      Height          =   165
      Left            =   750
      TabIndex        =   28
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person: "
      Height          =   165
      Left            =   750
      TabIndex        =   27
      Top             =   1530
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Debtor Number: "
      Height          =   195
      Left            =   660
      TabIndex        =   26
      Top             =   810
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Debtor Name: "
      Height          =   195
      Left            =   660
      TabIndex        =   25
      Top             =   1170
      Width           =   1305
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   1380
      X2              =   11010
      Y1              =   630
      Y2              =   630
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   35
      Left            =   60
      TabIndex        =   23
      Top             =   510
      Width           =   1785
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "General"
      Size            =   "3149;397"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image Image3 
      Height          =   1305
      Left            =   2010
      Top             =   3270
      Width           =   3315
      BorderColor     =   0
      BackColor       =   16777215
      Size            =   "5847;2302"
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Code: "
      Height          =   225
      Index           =   1
      Left            =   6780
      TabIndex        =   22
      Top             =   1890
      Width           =   1155
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   39
      Left            =   900
      TabIndex        =   24
      Top             =   180
      Width           =   3105
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Debtor Details"
      Size            =   "5477;397"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Left            =   840
      Top             =   150
      Width           =   3195
      BackColor       =   16777215
      Size            =   "5636;556"
      VariousPropertyBits=   19
   End
End
Attribute VB_Name = "frmGuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAccount_Click()
    Load frmAccount
    TillData.Account_No = grdSupp.TextMatrix(grdSupp.Row, 0)
    frmAccount.Tag = "Debtor"
    frmAccount.Show vbModal
End Sub

Private Sub cmdAll_Click()
    frmStateRun.Show vbModal
End Sub

Private Sub cmdChange_Click()
    frmDebtChange.Show vbModal
End Sub

Private Sub cmdDiscount_Click()
    frmDiscStruc.Show vbModal
End Sub

Private Sub cmdEmail_Click()
frmEmail.Show vbModal
End Sub

Private Sub cmdPricelist_Click()
If cmdPricelist.Tag = 1 Then
    cmdPricelist.Tag = 0
    Screen.MousePointer = 11
    Label5(2).Caption = "Account Details..."
    cmdStatement.Visible = True
    cmdAll.Visible = True
    cmdEmail.Visible = True
    cmdAccount.Visible = True
    Label12.Visible = True
    Label13.Visible = True
    DTStart.Visible = True
    DTStop.Visible = True
    cmdPrintpricelist.Visible = False
    If cmdPrintselected.Visible = True Then cmdPrintselected.Visible = False
    grdSupp.Cols = 5
    grdSupp.TextMatrix(0, 0) = "Debtor Number"
    grdSupp.TextMatrix(0, 1) = "Debtor Name"
    grdSupp.TextMatrix(0, 2) = "Contact Person"
    grdSupp.TextMatrix(0, 3) = "Tel.Number"
    grdSupp.TextMatrix(0, 4) = "Balance "
    grdSupp.ColWidth(0) = grdSupp.Width * 0.2
    grdSupp.ColWidth(1) = grdSupp.Width * 0.3
    grdSupp.ColWidth(2) = grdSupp.Width * 0.22
    grdSupp.ColWidth(3) = grdSupp.Width * 0.15
    grdSupp.ColWidth(4) = grdSupp.Width * 0.13
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    If grdSupp.Rows = 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
    Else
        grdSupp.Row = 1
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
    grdSupp.ColAlignment(0) = flexAlignLeftCenter
    grdSupp.ColAlignment(1) = flexAlignLeftCenter
    grdSupp.ColAlignment(2) = flexAlignLeftCenter
    grdSupp.ColAlignment(3) = flexAlignLeftCenter
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    LoadSuppliers
    frmMain.stbBar.Panels(3) = "Records = " & Val(grdSupp.Rows - 1)
    YYYY = Format(Date, "YYYY")
    MM = Format(Date, "MM")
    DTStart.Value = YYYY & "-" & MM & "-01"
    DTStop.Value = Date
    Screen.MousePointer = 0
    cmdPricelist.Caption = "This Debtor's Pricelist..."
    Pricelistpreview = "No"
    Exit Sub
End If

'pricelistcode
cmdPrintselected.Visible = True

ActiveUpdateServer "Delete from Pricelist_Temp"
Label5(2).Caption = "Account Pricelist..."
    cmdStatement.Visible = False
    cmdAll.Visible = False
    cmdEmail.Visible = False
    cmdAccount.Visible = False
    Label12.Visible = False
    Label13.Visible = False
    DTStart.Visible = False
    DTStop.Visible = False
    cmdPrintpricelist.Visible = False
grdSupp.Clear
   'grdSupp.Cols = 7
   '****************
   grdSupp.Cols = 8
    
    grdSupp.TextMatrix(0, 0) = " Product Code"
    grdSupp.TextMatrix(0, 1) = " Description"
    grdSupp.TextMatrix(0, 2) = "Unit"
    grdSupp.TextMatrix(0, 3) = "Measurement"
    grdSupp.TextMatrix(0, 4) = "Price2"
    grdSupp.TextMatrix(0, 5) = "Client Price"
    grdSupp.TextMatrix(0, 6) = "Normal Price"
    '****************
    grdSupp.TextMatrix(0, 7) = "Print"
    grdSupp.ColDataType(7) = flexDTBoolean
   
   
   grdSupp.ColWidth(0) = grdSupp.Width * 0.12
    grdSupp.ColWidth(1) = grdSupp.Width * 0.26
    grdSupp.ColWidth(2) = grdSupp.Width * 0.08
    grdSupp.ColWidth(3) = grdSupp.Width * 0.12
    grdSupp.ColWidth(4) = grdSupp.Width * 0.12
    grdSupp.ColWidth(5) = grdSupp.Width * 0.12
    grdSupp.ColWidth(6) = grdSupp.Width * 0.12
    '****************
    grdSupp.ColWidth(7) = grdSupp.Width * 0.016
    
    
    If grdSupp.Rows = 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
    Else
        grdSupp.Row = 1
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
    grdSupp.ColAlignment(0) = flexAlignLeftCenter
    grdSupp.ColAlignment(1) = flexAlignLeftCenter
    grdSupp.ColAlignment(2) = flexAlignLeftCenter
    grdSupp.ColAlignment(3) = flexAlignLeftCenter
    grdSupp.ColAlignment(4) = flexAlignCenterCenter
    grdSupp.ColAlignment(5) = flexAlignCenterCenter
    grdSupp.ColAlignment(6) = flexAlignCenterCenter
    '****************
    grdSupp.ColAlignment(7) = flexAlignCenterCenter
    
    
    Screen.MousePointer = 11
    DoEvents
    LoadPricelist
    
    frmMain.stbBar.Panels(3) = "Records = " & Val(grdSupp.Rows - 1)
    YYYY = Format(Date, "YYYY")
    MM = Format(Date, "MM")
    DTStart.Value = YYYY & "-" & MM & "-01"
    DTStop.Value = Date
    Screen.MousePointer = 0
    cmdPricelist.Tag = 1
    cmdPricelist.Caption = "Close Pricelist..."
    Pricelistpreview = "Yes"
End Sub
Private Sub LoadPricelist()
DoEvents
 Dim f As Integer

 grdSupp.Rows = 1
    
    grdSupp.Rows = grdSupp.Rows + 1
    ActiveReadServer2 "SELECT * From Full_Pricelist_view order by Description"
    If rs2.RecordCount > 0 Then
    ActiveReadServer "Select * from Debtor_Discounts where Debtor_No='" & txtSuppNo.Text & "'"
    If rs.RecordCount = 0 Then GoTo Cleanlist
    If rs.RecordCount > 0 Then
            rs.Close
            i = 0
            While Not rs2.EOF
checker:
                    If rs2.EOF Then GoTo Donemebit:
                    If rs2.Fields("Sales_Item") <> 1 Then rs2.MoveNext: GoTo checker
                    If rs2.Fields("Selling_Price") = 0 Then rs2.MoveNext: GoTo checker
                  
                    i = i + 1
            grdSupp.TextMatrix(i, 0) = rs2.Fields("Product_Code")
            grdSupp.TextMatrix(i, 1) = LTrim(rs2.Fields("Description"))
            grdSupp.TextMatrix(i, 2) = rs2.Fields("Unit_Size")
            grdSupp.TextMatrix(i, 3) = rs2.Fields("Unit_of_Measure")
            If rs2.Fields("Unit_Size") = 0 Then grdSupp.TextMatrix(i, 2) = ""
            grdSupp.TextMatrix(i, 4) = Format(rs2.Fields("Price2"), "00.00")
            grdSupp.TextMatrix(i, 6) = Format(rs2.Fields("Selling_Price"), "00.00")
            ActiveReadServer "Select * from Debtor_Discounts where Debtor_No='" & txtSuppNo.Text & "'"
                If rs.RecordCount > 0 Then
                rs.MoveFirst
                
                For f = 0 To rs.RecordCount - 1
                If rs.Fields("Department_No") = rs2.Fields("Department_No") And rs.Fields("Selling_Price") = 2 Then grdSupp.TextMatrix(i, 5) = Format(rs2.Fields("Price2"), "00.00")
               If rs.Fields("Department_No") = rs2.Fields("Department_No") And rs.Fields("Sell_Disc") <> 0 Then
               Dim x, y, z
               x = rs2.Fields("Selling_Price").Value
               y = ((rs2.Fields("Selling_Price").Value) * (rs.Fields("Sell_Disc").Value) / 100)
               z = x - y
               grdSupp.TextMatrix(i, 5) = Format(z, "00.00")

               End If
               If rs.Fields("Department_No") = rs2.Fields("Department_No") And UCase(rs.Fields("Sell_Cost")) = "YES" Then
               
               grdSupp.TextMatrix(i, 5) = Format(rs2.Fields("Landed_Cost").Value, "00.00")

               End If
               
               If rs.Fields("Department_No") = rs2.Fields("Department_No") And rs.Fields("Cost_Disc") <> 0 Then
                Dim xx, yy, zzz
               xx = rs2.Fields("Landed_Cost").Value
               yy = ((rs2.Fields("Landed_Cost").Value) * (rs.Fields("Cost_Disc").Value) / 100)
               zz = ((xx + yy) + ((14 / 100) * ((xx + yy))))
               grdSupp.TextMatrix(i, 5) = Format(zz, "00.00")
               End If
                
                If grdSupp.TextMatrix(i, 5) = "" Then grdSupp.TextMatrix(i, 5) = Format(rs2.Fields("Selling_Price"), "00.00")
                rs.MoveNext
                Next f
                
                rs2.MoveNext
                grdSupp.Rows = grdSupp.Rows + 1
                End If
            
        
        Wend
        End If
        

End If
 
    If rs2.State = 1 Then rs2.Close
    If rs.State = 1 Then rs.Close
   
     Updatetemppricelist
     Screen.MousePointer = vbNormal
     Exit Sub

 
Cleanlist:
 grdSupp.Rows = 1
        
        ActiveReadServer2 "SELECT * From Full_Pricelist_view order by Description"
        If rs2.RecordCount > 0 Then
            i = 0
            
            While Not rs2.EOF
checker2:
                    If rs2.EOF Then GoTo Donemebit:
                    If rs2.Fields("Sales_Item") <> 1 Then rs2.MoveNext: GoTo checker2
                    If rs2.Fields("Selling_Price") = 0 Then rs2.MoveNext: GoTo checker2
                 i = i + 1
                If rs2.EOF Then GoTo Donemebit
                grdSupp.Rows = grdSupp.Rows + 1
                grdSupp.TextMatrix(i, 0) = rs2.Fields("Product_Code")
                grdSupp.TextMatrix(i, 1) = LTrim(rs2.Fields("Description"))
                grdSupp.TextMatrix(i, 2) = rs2.Fields("Unit_Size")
                grdSupp.TextMatrix(i, 3) = rs2.Fields("Unit_of_Measure")
                If rs2.Fields("Unit_Size") = 0 Then grdSupp.TextMatrix(i, 2) = ""
                grdSupp.TextMatrix(i, 4) = Format(rs2.Fields("Price2"), "00.00")
                grdSupp.TextMatrix(i, 5) = Format(rs2.Fields("Selling_Price"), "00.00")
                grdSupp.TextMatrix(i, 6) = Format(rs2.Fields("Selling_Price"), "00.00")
                rs2.MoveNext
                
                Wend
            End If
    
   
Donemebit:
If rs2.State = 1 Then rs2.Close
If rs.State = 1 Then rs.Close
ActiveUpdateServer "Delete from Pricelist_Temp"
If rs.State = 1 Then rs.Close
Updatetemppricelist
End Sub

Private Sub Updatetemppricelist()
   Dim a As String
   Dim Unitval As String
  DoEvents
  With frmGuests
        For i = 1 To frmGuests.grdSupp.Rows - 1
            If grdSupp.TextMatrix(i, 0) = "" Then GoTo Ender
            a = Format(.grdSupp.TextMatrix(i, 5), "00.00")
            Unitval = .grdSupp.TextMatrix(i, 2)
            If .grdSupp.TextMatrix(i, 2) = "" Then Unitval = "Null"
            ActiveUpdateServer "INSERT INTO Pricelist_Temp (Product_Code, Description, Unit, Measurement, Price)" & _
            " VALUES ('" & .grdSupp.TextMatrix(i, 0) & "','" & .grdSupp.TextMatrix(i, 1) & "','" & Unitval & "','" & .grdSupp.TextMatrix(i, 3) & "','" & a & "')"
        Next i
    End With
                
Ender:
     If rs.State = 1 Then rs.Close

End Sub




Private Sub cmdPrintselected_Click()
 
 Dim a As String
   Dim Unitval As String
  DoEvents
  ActiveUpdateServer "Delete from Pricelist_Temp"
  
  For w = 1 To grdSupp.Rows - 1
 grdSupp.Row = w
  If grdSupp.ValueMatrix(w, 7) <> 0 Then
  With frmGuests
        
            If grdSupp.TextMatrix(w, 0) = "" Then GoTo Ender
            a = Format(.grdSupp.TextMatrix(w, 5), "00.00")
            Unitval = .grdSupp.TextMatrix(w, 2)
            If .grdSupp.TextMatrix(w, 2) = "" Then Unitval = "Null"
            ActiveUpdateServer "INSERT INTO Pricelist_Temp (Product_Code, Description, Unit, Measurement, Price)" & _
            " VALUES ('" & .grdSupp.TextMatrix(w, 0) & "','" & .grdSupp.TextMatrix(w, 1) & "','" & Unitval & "','" & .grdSupp.TextMatrix(w, 3) & "','" & a & "') "
   
    End With
    End If
     Next w
   DoEvents
        

Ender:
     If rs.State = 1 Then rs.Close
    
     rptPricelist.Show
     

End Sub

Private Sub cmdStatement_Click()
    If grdSupp.Row <> 0 Then
        TillData.Account_No = grdSupp.TextMatrix(grdSupp.Row, 0)
        rptStatement.Show
    End If
End Sub
Private Sub cmdUp_Click()
    Select Case cmdUp.Caption
        Case "5"
            grdSupp.SetFocus
            picHead.top = 0
            grdSupp.top = picHead.top + picHead.Height
            grdSupp.Height = frmGuests.Height - picHead.Height - 120
            DoEvents
            cmdUp.Caption = 6
        Case "6"
            cmdUp.Caption = "5"
            picHead.top = 4650
            grdSupp.top = 5430
            grdSupp.Height = 3405
            DoEvents
            txtSuppNo.SetFocus
    End Select
End Sub



Private Sub Form_Activate()
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Debtors"
    frmMain.picAccBar.Enabled = True
    If Emailsetup = False Then cmdEmail.Enabled = False
    SafetyCode "Debtors"
    
End Sub


Private Sub Form_Load()
cmdPricelist.Tag = 0
    Screen.MousePointer = 11
    grdSupp.Cols = 5
    grdSupp.TextMatrix(0, 0) = " Debtor Number"
    grdSupp.TextMatrix(0, 1) = " Debtor Name"
    grdSupp.TextMatrix(0, 2) = "Contact Person"
    grdSupp.TextMatrix(0, 3) = "Tel.Number"
    grdSupp.TextMatrix(0, 4) = "Balance "
    grdSupp.ColWidth(0) = grdSupp.Width * 0.2
    grdSupp.ColWidth(1) = grdSupp.Width * 0.3
    grdSupp.ColWidth(2) = grdSupp.Width * 0.22
    grdSupp.ColWidth(3) = grdSupp.Width * 0.15
    grdSupp.ColWidth(4) = grdSupp.Width * 0.13
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    If grdSupp.Rows = 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
    Else
        grdSupp.Row = 1
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
    grdSupp.ColAlignment(0) = flexAlignLeftCenter
    grdSupp.ColAlignment(1) = flexAlignLeftCenter
    grdSupp.ColAlignment(2) = flexAlignLeftCenter
    grdSupp.ColAlignment(3) = flexAlignLeftCenter
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    LoadSuppliers
    frmMain.stbBar.Panels(3) = "Records = " & Val(grdSupp.Rows - 1)
    YYYY = Format(Date, "YYYY")
    MM = Format(Date, "MM")
    DTStart.Value = YYYY & "-" & MM & "-01"
    DTStop.Value = Date
    Screen.MousePointer = 0
End Sub
Public Sub LoadSuppliers()
    If cmdPricelist.Tag = 1 Then
          Cleanlist
            End If
    txtSuppNo.Text = ""
    txtSuppName.Text = ""
    txtContact.Text = ""
    txtContact.Text = ""
    txtCredit.Text = ""
    txtBussTell.Text = ""
    txtCell.Text = ""
    txtFax.Text = ""
    txtAddress.Text = ""
    txtEmail.Text = ""
    txtWeb.Text = ""
    txtGL_Code.Text = ""
    optTerms(0).Value = True
     Screen.MousePointer = 11
    grdSupp.Cols = 5
    grdSupp.TextMatrix(0, 0) = " Debtor Number"
    grdSupp.TextMatrix(0, 1) = " Debtor Name"
    grdSupp.TextMatrix(0, 2) = "Contact Person"
    grdSupp.TextMatrix(0, 3) = "Tel.Number"
    grdSupp.TextMatrix(0, 4) = "Balance "
    grdSupp.ColWidth(0) = grdSupp.Width * 0.2
    grdSupp.ColWidth(1) = grdSupp.Width * 0.3
    grdSupp.ColWidth(2) = grdSupp.Width * 0.22
    grdSupp.ColWidth(3) = grdSupp.Width * 0.15
    grdSupp.ColWidth(4) = grdSupp.Width * 0.13
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    If grdSupp.Rows = 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
    Else
        grdSupp.Row = 1
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
    grdSupp.ColAlignment(0) = flexAlignLeftCenter
    grdSupp.ColAlignment(1) = flexAlignLeftCenter
    grdSupp.ColAlignment(2) = flexAlignLeftCenter
    grdSupp.ColAlignment(3) = flexAlignLeftCenter
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    'LoadSuppliers
    frmMain.stbBar.Panels(3) = "Records = " & Val(grdSupp.Rows - 1)
    YYYY = Format(Date, "YYYY")
    MM = Format(Date, "MM")
    DTStart.Value = YYYY & "-" & MM & "-01"
    DTStop.Value = Date
    Screen.MousePointer = 0
    
    grdSupp.Rows = 1
    If frmMain.cmdMenu(11).Value <> 0 Then
        DType = 0
    End If
    If frmMain.cmdMenu(9).Value <> 0 Then
        DType = 1
    End If
    If frmMain.cmdMenu(10).Value <> 0 Then
        DType = 2
    End If
    If frmMain.cmdMenu(12).Value <> 0 Then
        DType = 3
    End If
    If frmMain.cmdMenu(14).Value <> 0 Then
        DType = 4
    End If
    ActiveReadServer "Select * from Debtors where Debt_Type = " & DType & " order by Debtor_No"
    grdSupp.Rows = rs.RecordCount + 1
    i = 0
    While Not rs.EOF
        i = i + 1
        grdSupp.TextMatrix(i, 0) = rs.Fields("Debtor_No")
        grdSupp.TextMatrix(i, 1) = rs.Fields("Debtor_Name")
        grdSupp.TextMatrix(i, 2) = rs.Fields("Contact_Person") & ""
        grdSupp.TextMatrix(i, 3) = rs.Fields("Business_Tel") & ""
        grdSupp.TextMatrix(i, 4) = Format(rs.Fields("Balance"), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    If grdSupp.Rows > 1 Then grdSupp.Row = 1
    
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
        cmdDiscount.Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
        cmdDiscount.Enabled = False
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
    frmMain.picAccBar.Enabled = False
End Sub

Private Sub grdSupp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Label5(2).Caption = "Account Pricelist..." Then Exit Sub
    On Error Resume Next
    If grdSupp.Rows = 1 Or grdSupp.Tag = "1" Then Exit Sub
    If OldRow <> NewRow Then
        ActiveReadServer "Select * from Debtors where Debtor_No='" & grdSupp.TextMatrix(grdSupp.Row, 0) & "'"
        If rs.RecordCount > 0 Then
            txtSuppNo = rs.Fields("Debtor_No")
            txtSuppName = rs.Fields("Debtor_Name")
            txtContact = rs.Fields("Contact_Person")
            txtAddress = rs.Fields("Address")
            txtBussTell = rs.Fields("Business_Tel")
            txtCell = rs.Fields("Mobile")
            txtEmail = rs.Fields("E_Mail")
            txtCredit = Format(rs.Fields("Credit_Limit"), "0.00")
            txtFax = rs.Fields("Fax_Tel")
            txtWeb = rs.Fields("Web_Page") & ""
            ActiveReadServer1 "Select * from Debtor_Discounts where Debtor_No = '" & rs.Fields("Debtor_No") & "'"
            If rs1.RecordCount > 0 Then
                txtVat.Text = "Set Per Department"
            Else
                txtVat.Text = "<Not Set>"
            End If
            rs1.Close
            txtGL_Code.Text = rs.Fields("GL_Code") & ""
            For i = 0 To optTerms.Count - 1
                If i = Val(rs.Fields("Terms") & "") Then
                    optTerms(i).Value = True
                    Exit For
                End If
            Next i
            frmMain.Toolbar1.Buttons(5).Enabled = True
        End If
        rs.Close
        frmMain.Toolbar1.Buttons(2).Enabled = True
    End If
    On Error GoTo 0
End Sub
Private Sub grdSupp_AfterSort(ByVal Col As Long, Order As Integer)
    On Error Resume Next
    If grdSupp.Rows = 1 Or grdSupp.Tag = "1" Then Exit Sub
    ActiveReadServer "Select * from Debtors where Debtor_No='" & grdSupp.TextMatrix(grdSupp.Row, 0) & "'"
    If rs.RecordCount > 0 Then
        txtSuppNo = rs.Fields("Debtor_No")
        txtSuppName = rs.Fields("Debtor_Name")
        txtContact = rs.Fields("Contact_Person")
        txtAddress = rs.Fields("Address")
        txtBussTell = rs.Fields("Business_Tel")
        txtCell = rs.Fields("Mobile")
        txtEmail = rs.Fields("E_Mail")
        txtCredit = Format(rs.Fields("Credit_Limit"), "0.00")
        txtFax = rs.Fields("Fax_Tel")
        txtWeb = rs.Fields("Web_Page")
        ActiveReadServer1 "Select * from Debtor_Discounts where Debtor_No = '" & rs.Fields("Debtor_No") & "'"
        If rs1.RecordCount > 0 Then
            txtVat.Text = "Set Per Department"
        Else
            txtVat.Text = "<Not Set>"
        End If
        rs1.Close
        txtGL_Code.Text = rs.Fields("GL_Code") & ""
        For i = 0 To optTerms.Count - 1
            If i = Val(rs.Fields("Terms") & "") Then
                optTerms(i).Value = True
                Exit For
            End If
        Next i
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
    rs.Close
    frmMain.Toolbar1.Buttons(2).Enabled = True
    On Error GoTo 0
End Sub

Private Sub grdSupp_Click()
If grdSupp.Col = 7 Then
   
        grdSupp.Editable = flexEDKbdMouse
    Else
        grdSupp.Editable = flexEDNone
    End If
End Sub

Private Sub optTerms_Click(Index As Integer)
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub



Private Sub timerwindow_Timer()
frmGuests.WindowState = 2
timerwindow.Enabled = False
End Sub

Private Sub txtAddress_Change()
On Error Resume Next
    Rows = 0
    For i = 1 To Len(txtAddress.Text)
        If Asc(Mid(txtAddress.Text, i, 1)) = 13 Then
            Rows = Rows + 1
            If Rows = 6 Then
                cmbRegion.SetFocus
                Exit For
            End If
        End If
    Next i
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
On Error GoTo 0
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtAddress_LostFocus()
    newstring = txtAddress.Text
    For i = 1 To Len(txtAddress.Text)
        If i = 1 Then
            If Asc(Mid(txtAddress.Text, 1, 1)) > 96 And Asc(Mid(txtAddress.Text, 1, 1)) < 123 Then
                Mid(newstring, i, 1) = UCase(Mid(txtAddress.Text, 1, 1))
            End If
        End If
        If Asc(Mid(txtAddress.Text, i, 1)) = 13 Then
            If i + 1 = Len(txtAddress.Text) Then
            Else
                Mid(newstring, i + 2, 1) = UCase(Mid(txtAddress.Text, i + 2, 1))
            End If
        End If
    Next i
    txtAddress.Text = Trim(newstring)
End Sub

Private Sub txtBussTell_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtBussTell_GotFocus()
    txtBussTell.SelStart = 0
    txtBussTell.SelLength = Len(txtBussTell.Text)
End Sub

Private Sub txtBussTell_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtCredit.SetFocus
        Case 40: KeyCode = 0: txtCell.SetFocus
    End Select
End Sub

Private Sub txtBussTell_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 45, 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCell_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtCell_GotFocus()
    txtCell.SelStart = 0
    txtCell.SelLength = Len(txtCell.Text)
End Sub

Private Sub txtCell_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtBussTell.SetFocus
        Case 40: KeyCode = 0: txtFax.SetFocus
    End Select
End Sub

Private Sub txtCell_KeyPress(KeyAscii As Integer)
         Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 45, 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtContact_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtContact_GotFocus()
    txtContact.SelStart = 0
    txtContact.SelLength = Len(txtContact.Text)
End Sub

Private Sub txtContact_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtSuppName.SetFocus
        Case 40: KeyCode = 0: txtCredit.SetFocus
    End Select
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtContact_LostFocus()
    On Error Resume Next
    txtContact.Text = UCase(Left(txtContact.Text, 1)) & Mid(txtContact.Text, 2)
    On Error GoTo 0
End Sub

Private Sub txtCredit_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtCredit_GotFocus()
    txtCredit.SelStart = 0
    txtCredit.SelLength = Len(txtCredit.Text)
End Sub

Private Sub txtCredit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtContact.SetFocus
        Case 40: KeyCode = 0: txtBussTell.SetFocus
    End Select
End Sub

Private Sub txtCredit_KeyPress(KeyAscii As Integer)
    If InStr(txtCredit.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 32
            KeyAscii = 0
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtCredit_LostFocus()
    If Val(txtCredit.Text) = 0 Then txtCredit.Text = "0.00"
    txtCredit.Text = Format(txtCredit.Text, "0.00")
End Sub
Private Sub txtEmail_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtEmail_GotFocus()
    txtEmail.SelStart = 0
    txtEmail.SelLength = Len(txtEmail.Text)
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtVat.SetFocus
        Case 40: KeyCode = 0: txtWeb.SetFocus
    End Select
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 95
        Exit Sub
        Case 8
        Case 39
            KeyAscii = 0
        Case 97 To 122
        Case 45, 46, 48 To 57, 64 To 90
        Case Else
            KeyAscii = 0

    End Select
End Sub

Private Sub txtFax_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtFax_GotFocus()
    txtFax.SelStart = 0
    txtFax.SelLength = Len(txtFax.Text)
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtCell.SetFocus
        Case 40: KeyCode = 0: txtAddress.SetFocus
    End Select
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
         Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 45, 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtGL_Code_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 97 To 122
        Case 45, 46, 48 To 57, 64 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSuppName_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtSuppName_GotFocus()
    txtSuppName.SelStart = 0
    txtSuppName.SelLength = Len(txtSuppName.Text)
End Sub

Private Sub txtSuppName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtSuppNo.SetFocus
        Case 40: KeyCode = 0: txtContact.SetFocus
    End Select
End Sub

Private Sub txtSuppName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtSuppName_LostFocus()
    On Error Resume Next
    txtSuppName.Text = UCase(Left(txtSuppName.Text, 1)) & Mid(txtSuppName.Text, 2)
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        cmdDiscount.Enabled = True
    Else
        cmdDiscount.Enabled = False
    End If
    On Error GoTo 0
End Sub
Private Sub txtSuppNo_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
        cmdDiscount.Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
        cmdDiscount.Enabled = False
    End If
End Sub
Private Sub txtSuppNo_GotFocus()
    txtSuppNo.SelStart = 0
    txtSuppNo.SelLength = Len(txtSuppNo.Text)
End Sub

Private Sub txtSuppNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            KeyCode = 0
            txtSuppName.SetFocus
            DoEvents
        Case 38
            KeyCode = 0
            txtWeb.SetFocus
        Case 40
            KeyCode = 0
            txtSuppName.SetFocus
    End Select
End Sub
Private Sub txtSuppNo_KeyPress(KeyAscii As Integer)
    If txtSuppName.Text <> "" Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 32
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSuppNo_LostFocus()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        cmdDiscount.Enabled = True
    Else
        cmdDiscount.Enabled = False
    End If
    If txtSuppName.Text = "" Then
        ActiveReadServer "Select * from Debtors where Debtor_No = '" & txtSuppNo.Text & "'"
        If rs.RecordCount > 0 Then
            MsgBox "You already have a Debtor with this number listed.", vbCritical, "HeroPOS"
            txtSuppNo.Text = ""
            txtSuppNo.SetFocus
            Exit Sub
        End If
        rs.Close
    End If
End Sub
Private Sub txtVat_GotFocus()
    txtVat.SelStart = 0
    txtVat.SelLength = Len(txtVat.Text)
End Sub

Private Sub txtVAT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtAddress.SetFocus
        Case 40: KeyCode = 0: txtEmail.SetFocus
    End Select
End Sub

Private Sub txtVat_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtWeb_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtWeb_GotFocus()
    txtWeb.SelStart = 0
    txtWeb.SelLength = Len(txtWeb.Text)
End Sub

Private Sub txtWeb_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtEmail.SetFocus
        Case 40: KeyCode = 0: optTerms(0).SetFocus
    End Select
End Sub

Private Sub txtWeb_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 97 To 122
        Case 45, 46, 48 To 57, 64 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub
Public Sub Cleanlist()
If cmdPricelist.Tag = 1 Then
    cmdPricelist.Tag = 0
    Screen.MousePointer = 11
    Label5(2).Caption = "Account Details..."
    cmdStatement.Visible = True
    cmdAll.Visible = True
    cmdEmail.Visible = True
    cmdAccount.Visible = True
    Label12.Visible = True
    Label13.Visible = True
    DTStart.Visible = True
    DTStop.Visible = True
    If cmdPrintselected.Visible = True Then cmdPrintselected.Visible = False
    
    grdSupp.Cols = 5
    grdSupp.TextMatrix(0, 0) = " Debtor Number"
    grdSupp.TextMatrix(0, 1) = " Debtor Name"
    grdSupp.TextMatrix(0, 2) = "Contact Person"
    grdSupp.TextMatrix(0, 3) = "Tel.Number"
    grdSupp.TextMatrix(0, 4) = "Balance "
    grdSupp.ColWidth(0) = grdSupp.Width * 0.2
    grdSupp.ColWidth(1) = grdSupp.Width * 0.3
    grdSupp.ColWidth(2) = grdSupp.Width * 0.22
    grdSupp.ColWidth(3) = grdSupp.Width * 0.15
    grdSupp.ColWidth(4) = grdSupp.Width * 0.13
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    If grdSupp.Rows = 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
    Else
        grdSupp.Row = 1
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
    grdSupp.ColAlignment(0) = flexAlignLeftCenter
    grdSupp.ColAlignment(1) = flexAlignLeftCenter
    grdSupp.ColAlignment(2) = flexAlignLeftCenter
    grdSupp.ColAlignment(3) = flexAlignLeftCenter
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    LoadSuppliers
    frmMain.stbBar.Panels(3) = "Records = " & Val(grdSupp.Rows - 1)
    YYYY = Format(Date, "YYYY")
    MM = Format(Date, "MM")
    DTStart.Value = YYYY & "-" & MM & "-01"
    DTStop.Value = Date
    Screen.MousePointer = 0
    cmdPricelist.Caption = "This Debtor's Pricelist..."
    Exit Sub
    End If


End Sub
