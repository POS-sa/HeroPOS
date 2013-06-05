VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSuppliers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MousePointer    =   4  'Icon
   ScaleHeight     =   9630
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid grdSupp 
      Height          =   3405
      Left            =   -30
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
      FormatString    =   $"frmSuppliers.frx":0000
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
   Begin VB.CheckBox chkLand 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GRV's must not Change the Landed Cost Price."
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7050
      TabIndex        =   38
      Top             =   3960
      Width           =   2925
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   -30
      ScaleHeight     =   765
      ScaleWidth      =   14865
      TabIndex        =   32
      Top             =   4650
      Width           =   14865
      Begin VB.PictureBox picLinks 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   8850
         ScaleHeight     =   585
         ScaleWidth      =   4215
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
         Begin btButtonEx.ButtonEx cmdLinker 
            Height          =   465
            Left            =   3450
            TabIndex        =   43
            Top             =   60
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   820
            Appearance      =   3
            Caption         =   "Link"
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
         Begin MSForms.ComboBox cmbSupp 
            Height          =   480
            Left            =   30
            TabIndex        =   42
            Top             =   60
            Width           =   3375
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "5953;847"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Arial Narrow"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin btButtonEx.ButtonEx cmdUp 
         Height          =   465
         Left            =   13110
         TabIndex        =   18
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
      Begin btButtonEx.ButtonEx cmdAccount 
         Height          =   465
         Left            =   11010
         TabIndex        =   37
         Top             =   180
         Width           =   2055
         _ExtentX        =   3625
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
      Begin btButtonEx.ButtonEx cmdStatement 
         Height          =   465
         Left            =   8880
         TabIndex        =   39
         Top             =   180
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "Print Payment Advice"
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
      Begin btButtonEx.ButtonEx cmdLinks 
         Height          =   465
         Left            =   7050
         TabIndex        =   40
         Top             =   180
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   820
         Appearance      =   3
         Caption         =   "Show Product Links"
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
         Left            =   5130
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   135
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   67895299
         CurrentDate     =   38862
      End
      Begin MSComCtl2.DTPicker DTStop 
         Height          =   285
         Left            =   5130
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   405
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   67895299
         CurrentDate     =   38862
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Advice from:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3390
         TabIndex        =   47
         Top             =   195
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Advice to:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3390
         TabIndex        =   46
         Top             =   480
         Width           =   1695
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
      Begin VB.Label Label5 
         Caption         =   "Supplier Details..."
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
         Left            =   240
         TabIndex        =   33
         Top             =   210
         Width           =   3135
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   2
         Left            =   3270
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   1
         Left            =   2940
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   0
         Left            =   2610
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image6 
         Height          =   90
         Left            =   150
         Top             =   570
         Width           =   2415
         BackColor       =   16761024
         Size            =   "4260;159"
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
   Begin VB.Frame Terms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Terms"
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   7050
      TabIndex        =   19
      Top             =   2310
      Width           =   2895
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
         Caption         =   "120 Days +"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1710
         TabIndex        =   16
         Top             =   1110
         Width           =   1155
      End
   End
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   1125
      Left            =   2130
      TabIndex        =   7
      Top             =   3360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1984
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmSuppliers.frx":0078
   End
   Begin MSForms.TextBox txtGL_Code 
      Height          =   285
      Left            =   7050
      TabIndex        =   36
      Top             =   1950
      Width           =   2895
      VariousPropertyBits=   746604563
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5106;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Code: "
      Height          =   225
      Index           =   1
      Left            =   5880
      TabIndex        =   35
      Top             =   2010
      Width           =   1155
   End
   Begin MSForms.TextBox txtVAT 
      Height          =   285
      Left            =   7050
      TabIndex        =   8
      Top             =   960
      Width           =   2895
      VariousPropertyBits=   746604563
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5106;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VAT Number: "
      Height          =   165
      Left            =   5880
      TabIndex        =   34
      Top             =   1020
      Width           =   1155
   End
   Begin MSForms.Image Image3 
      Height          =   1305
      Left            =   1980
      Top             =   3270
      Width           =   3315
      BorderColor     =   0
      BackColor       =   16777215
      Size            =   "5847;2302"
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Left            =   810
      Top             =   240
      Width           =   3195
      BackColor       =   16777215
      Size            =   "5636;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   35
      Left            =   30
      TabIndex        =   31
      Top             =   690
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
   Begin VB.Line Line1 
      Index           =   6
      X1              =   1350
      X2              =   9930
      Y1              =   780
      Y2              =   780
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   39
      Left            =   870
      TabIndex        =   30
      Top             =   270
      Width           =   3105
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Supplier Details"
      Size            =   "5477;397"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtSuppNo 
      Height          =   285
      Left            =   1980
      TabIndex        =   0
      Top             =   960
      Width           =   2355
      VariousPropertyBits=   746604563
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "4154;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtSuppName 
      Height          =   285
      Left            =   1980
      TabIndex        =   1
      Top             =   1290
      Width           =   3585
      VariousPropertyBits=   746604563
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "6324;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name: "
      Height          =   195
      Left            =   630
      TabIndex        =   29
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Number: "
      Height          =   195
      Left            =   630
      TabIndex        =   28
      Top             =   990
      Width           =   1305
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person: "
      Height          =   165
      Left            =   720
      TabIndex        =   27
      Top             =   1650
      Width           =   1215
   End
   Begin MSForms.TextBox txtContact 
      Height          =   285
      Left            =   1980
      TabIndex        =   2
      Top             =   1620
      Width           =   3585
      VariousPropertyBits=   746604563
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "6324;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business Tel: "
      Height          =   165
      Left            =   720
      TabIndex        =   25
      Top             =   2310
      Width           =   1215
   End
   Begin MSForms.TextBox txtBussTell 
      Height          =   285
      Left            =   1980
      TabIndex        =   4
      Top             =   2280
      Width           =   2745
      VariousPropertyBits=   746604563
      MaxLength       =   25
      BorderStyle     =   1
      Size            =   "4842;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile: "
      Height          =   165
      Left            =   720
      TabIndex        =   24
      Top             =   2640
      Width           =   1215
   End
   Begin MSForms.TextBox txtCell 
      Height          =   285
      Left            =   1980
      TabIndex        =   5
      Top             =   2610
      Width           =   2745
      VariousPropertyBits=   746604563
      MaxLength       =   25
      BorderStyle     =   1
      Size            =   "4842;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number: "
      Height          =   165
      Left            =   720
      TabIndex        =   23
      Top             =   2970
      Width           =   1215
   End
   Begin MSForms.TextBox txtFax 
      Height          =   285
      Left            =   1980
      TabIndex        =   6
      Top             =   2940
      Width           =   2745
      VariousPropertyBits=   746604563
      MaxLength       =   25
      BorderStyle     =   1
      Size            =   "4842;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address: "
      Height          =   165
      Left            =   720
      TabIndex        =   22
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: "
      Height          =   165
      Left            =   5880
      TabIndex        =   21
      Top             =   1350
      Width           =   1155
   End
   Begin MSForms.TextBox txtEmail 
      Height          =   285
      Left            =   7050
      TabIndex        =   9
      Top             =   1290
      Width           =   2895
      VariousPropertyBits=   746604563
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5106;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Web Page: "
      Height          =   225
      Index           =   0
      Left            =   5880
      TabIndex        =   20
      Top             =   1680
      Width           =   1155
   End
   Begin MSForms.TextBox txtWeb 
      Height          =   285
      Left            =   7050
      TabIndex        =   10
      Top             =   1620
      Width           =   2895
      VariousPropertyBits=   746604563
      MaxLength       =   30
      BorderStyle     =   1
      Size            =   "5106;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtCredit 
      Height          =   285
      Left            =   1980
      TabIndex        =   3
      Top             =   1950
      Width           =   1455
      VariousPropertyBits=   746604563
      MaxLength       =   12
      BorderStyle     =   1
      Size            =   "2566;503"
      Value           =   "0.00"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit: "
      Height          =   165
      Index           =   0
      Left            =   720
      TabIndex        =   26
      Top             =   1980
      Width           =   1215
   End
End
Attribute VB_Name = "frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbSupp_Change()
    If cmbSupp.Tag = "1" Then Exit Sub
    LoadSupplierLinks
End Sub

Private Sub cmdAccount_Click()
    Load frmAccount
    TillData.Account_No = grdSupp.TextMatrix(grdSupp.Row, 0)
    frmAccount.Tag = "Supplier"
    frmAccount.Show vbModal
    SaveRow = grdSupp.Row
    LoadSuppliers
    grdSupp.Row = SaveRow
End Sub

Private Sub cmdLinker_Click()
    Load frmSuppCodes
    frmSuppCodes.Tag = "Supp"
    frmSuppCodes.Show vbModal
    frmSuppCodes.Tag = ""
End Sub
Private Sub cmdLinks_Click()
    Select Case cmdLinks.Caption
        Case "Show Product Links"
            Screen.MousePointer = 11
            cmbSupp.Tag = "1"
            cmdLinks.Caption = "Hide Product Links"
            cmbSupp.Clear
            cmbSupp.AddItem "<Not Linked>"
            ActiveReadServer "Select Supplier_No,Supplier_Name from suppliers order by Supplier_name"
            cmbSupp.AddItem "<All Suppliers>"
            While Not rs.EOF
                cmbSupp.AddItem rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
                rs.MoveNext
            Wend
            rs.Close
            cmbSupp.Text = "<All Suppliers>"
            cmbSupp.Tag = ""
            picLinks.Visible = True
            grdSupp.Rows = 1
            grdSupp.Cols = 7
            grdSupp.TextMatrix(0, 0) = "Product Code"
            grdSupp.TextMatrix(0, 1) = "Description "
            grdSupp.TextMatrix(0, 2) = " Supplier Number"
            grdSupp.TextMatrix(0, 3) = "Supplier Name"
            grdSupp.TextMatrix(0, 4) = "Supplier Code"
            grdSupp.TextMatrix(0, 5) = "Landed Cost "
            grdSupp.TextMatrix(0, 6) = "List Price "
            grdSupp.ColWidth(0) = grdSupp.Width * 0.1
            grdSupp.ColWidth(1) = grdSupp.Width * 0.3
            grdSupp.ColWidth(2) = grdSupp.Width * 0.1
            grdSupp.ColWidth(3) = grdSupp.Width * 0.2
            grdSupp.ColWidth(4) = grdSupp.Width * 0.1
            grdSupp.ColWidth(5) = grdSupp.Width * 0.1
            grdSupp.ColWidth(6) = grdSupp.Width * 0.1
            frmMain.Toolbar1.Buttons(2).Enabled = False
            frmMain.Toolbar1.Buttons(4).Enabled = False
            frmMain.Toolbar1.Buttons(5).Enabled = False
            grdSupp.ColAlignment(0) = flexAlignLeftCenter
            grdSupp.ColAlignment(1) = flexAlignLeftCenter
            grdSupp.ColAlignment(2) = flexAlignLeftCenter
            grdSupp.ColAlignment(3) = flexAlignLeftCenter
            grdSupp.ColAlignment(4) = flexAlignLeftCenter
            grdSupp.ColAlignment(5) = flexAlignRightCenter
            grdSupp.ColAlignment(6) = flexAlignRightCenter
            LoadSupplierLinks
            frmMain.stbBar.Panels(3) = "Records = " & Val(grdSupp.Rows - 1)
            Screen.MousePointer = 0
        Case "Hide Product Links"
            Screen.MousePointer = 11
            cmdLinks.Caption = "Show Product Links"
            grdSupp.Cols = 5
            grdSupp.TextMatrix(0, 0) = " Supplier Number"
            grdSupp.TextMatrix(0, 1) = "Supplier Name"
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
            picLinks.Visible = False
            Screen.MousePointer = 0
    End Select
End Sub
Private Sub cmdStatement_Click()
    On Error Resume Next
    If grdSupp.Row <> 0 Then
        TillData.Account_No = grdSupp.TextMatrix(grdSupp.Row, 0)
        rptSuppStatement.Show
    End If
    On Error GoTo 0
End Sub

Private Sub cmdUp_Click()
    Select Case cmdUp.Caption
        Case "5"
            grdSupp.SetFocus
            picHead.top = 0
            grdSupp.top = picHead.top + picHead.Height
            grdSupp.Height = frmSuppliers.Height - picHead.Height - 120
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
    frmMain.Toolbar1.Buttons(16).Tag = "Suppliers"
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 11
    grdSupp.Cols = 5
    grdSupp.TextMatrix(0, 0) = " Supplier Number"
    grdSupp.TextMatrix(0, 1) = "Supplier Name"
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
    SafetyCode "Suppliers"
    YYYY = Format(Date, "YYYY")
    MM = Format(Date, "MM")
    DTStart.Value = YYYY & "-" & MM & "-01"
    DTStop.Value = Date
    Screen.MousePointer = 0
End Sub
Private Sub LoadSuppliers()
    grdSupp.Rows = 1
    ActiveReadServer "Select * from Suppliers order by Supplier_Name"
    grdSupp.Rows = rs.RecordCount + 1
    i = 0
    While Not rs.EOF
        i = i + 1
        grdSupp.TextMatrix(i, 0) = rs.Fields("Supplier_No")
        grdSupp.TextMatrix(i, 1) = rs.Fields("Supplier_Name")
        grdSupp.TextMatrix(i, 2) = rs.Fields("Contact_Person")
        grdSupp.TextMatrix(i, 3) = rs.Fields("Business_Tel")
        grdSupp.TextMatrix(i, 4) = Format(rs.Fields("Balance"), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    If grdSupp.Rows > 1 Then grdSupp.Row = 1
End Sub
Private Sub LoadSupplierLinks()
    grdSupp.Rows = 1
    If cmbSupp.Text = "<All Suppliers>" Then
        ActiveReadServer3 "Select * from Suppliers_Links_View order by Description"
    End If
    If cmbSupp.Text = "<Not Linked>" Then
        ActiveReadServer3 "Select * from Suppliers_Links_View where Supplier_No is null order by Description"
    End If
    If InStr(cmbSupp.Text, "-") <> 0 Then
        ActiveReadServer3 "Select * from Suppliers_Links_View where Supplier_No = '" & Trim(Mid(cmbSupp.Text, InStr(cmbSupp.Text, "-") + 1)) & "' order by Description"
    End If
    While Not rs3.EOF
        grdSupp.Rows = grdSupp.Rows + 1
        grdSupp.TextMatrix(grdSupp.Rows - 1, 0) = rs3.Fields("Product_Code")
        grdSupp.TextMatrix(grdSupp.Rows - 1, 1) = rs3.Fields("Description")
        grdSupp.TextMatrix(grdSupp.Rows - 1, 2) = rs3.Fields("Supplier_No") & ""
        grdSupp.TextMatrix(grdSupp.Rows - 1, 4) = rs3.Fields("Supplier_Code") & ""
        grdSupp.TextMatrix(grdSupp.Rows - 1, 3) = rs3.Fields("Supplier_Name") & ""
        grdSupp.TextMatrix(grdSupp.Rows - 1, 5) = Format(Val(rs3.Fields("List_Price") & ""), "0.00")
        grdSupp.TextMatrix(grdSupp.Rows - 1, 6) = Format(Val(rs3.Fields("Landed_Cost") & ""), "0.00")
        rs3.MoveNext
    Wend
    rs3.Close
    If grdSupp.Rows > 1 Then grdSupp.Row = 1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
End Sub
Private Sub grdSupp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If grdSupp.Rows = 1 Or grdSupp.Tag = "1" Then Exit Sub
    If OldRow <> NewRow Then
        ActiveReadServer "Select * from Suppliers where Supplier_No='" & grdSupp.TextMatrix(grdSupp.Row, 0) & "'"
        If rs.RecordCount > 0 Then
            txtSuppNo = rs.Fields("Supplier_No")
            txtSuppName = rs.Fields("Supplier_Name")
            txtContact = rs.Fields("Contact_Person")
            txtAddress = rs.Fields("Address")
            txtBussTell = rs.Fields("Business_Tel")
            txtCell = rs.Fields("Mobile")
            txtEmail = rs.Fields("E_Mail")
            txtCredit = Format(rs.Fields("Credit_Limit"), "0.00")
            txtFax = rs.Fields("Fax_Tel")
            txtWeb = rs.Fields("Web_Page") & ""
            txtVat.Text = rs.Fields("VAT_No") & ""
            txtGL_Code.Text = rs.Fields("GL_Code") & ""
            For i = 0 To optTerms.Count - 1
                If i = Val(rs.Fields("Terms") & "") Then
                    optTerms(i).Value = True
                    Exit For
                End If
            Next i
            chkLand.Value = Val(rs.Fields("Landed_Cost") & "")
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
    ActiveReadServer "Select * from Suppliers where Supplier_No='" & grdSupp.TextMatrix(grdSupp.Row, 0) & "'"
    If rs.RecordCount > 0 Then
        txtSuppNo = rs.Fields("Supplier_No")
        txtSuppName = rs.Fields("Supplier_Name")
        txtContact = rs.Fields("Contact_Person")
        txtAddress = rs.Fields("Address")
        txtBussTell = rs.Fields("Business_Tel")
        txtCell = rs.Fields("Mobile")
        txtEmail = rs.Fields("E_Mail")
        txtCredit = Format(rs.Fields("Credit_Limit"), "0.00")
        txtFax = rs.Fields("Fax_Tel")
        txtWeb = rs.Fields("Web_Page")
        txtVat.Text = rs.Fields("VAT_No") & ""
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

Private Sub optTerms_Click(Index As Integer)
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
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

Private Sub txtBussTell_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtCredit.SetFocus
        Case 40: KeyCode = 0: txtCell.SetFocus
    End Select
End Sub

Private Sub txtBussTell_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtCell_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtBussTell.SetFocus
        Case 40: KeyCode = 0: txtFax.SetFocus
    End Select
End Sub

Private Sub txtCell_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtContact_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtSuppName.SetFocus
        Case 40: KeyCode = 0: txtCredit.SetFocus
    End Select
End Sub

Private Sub txtContact_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtCredit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtContact.SetFocus
        Case 40: KeyCode = 0: txtBussTell.SetFocus
    End Select
End Sub

Private Sub txtCredit_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtEmail_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtVat.SetFocus
        Case 40: KeyCode = 0: txtWeb.SetFocus
    End Select
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtFax_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtCell.SetFocus
        Case 40: KeyCode = 0: txtAddress.SetFocus
    End Select
End Sub

Private Sub txtFax_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtGL_Code_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtSuppName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtSuppNo.SetFocus
        Case 40: KeyCode = 0: txtContact.SetFocus
    End Select
End Sub

Private Sub txtSuppName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtSuppName_LostFocus()
    On Error Resume Next
    txtSuppName.Text = UCase(Left(txtSuppName.Text, 1)) & Mid(txtSuppName.Text, 2)
    On Error GoTo 0
End Sub
Private Sub txtSuppNo_Change()
    If txtSuppNo.Text <> "" And txtSuppName.Text <> "" Then
        frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
    If grdSupp.FindRow(txtSuppNo.Text, 0, 0, 0, 1) > 0 Then
        grdSupp.Row = grdSupp.FindRow(txtSuppNo.Text, 0, 0, 0, 1)
        grdSupp.ShowCell grdSupp.Row, 0
    End If
End Sub
Private Sub txtSuppNo_GotFocus()
    txtSuppNo.SelStart = 0
    txtSuppNo.SelLength = Len(txtSuppNo.Text)
End Sub

Private Sub txtSuppNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtWeb.SetFocus
        Case 40: KeyCode = 0: txtSuppName.SetFocus
    End Select
End Sub

Private Sub txtSuppNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtVat_GotFocus()
    txtVat.SelStart = 0
    txtVat.SelLength = Len(txtVat.Text)
End Sub

Private Sub txtVAT_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtAddress.SetFocus
        Case 40: KeyCode = 0: txtEmail.SetFocus
    End Select
End Sub

Private Sub txtVat_KeyPress(KeyAscii As MSForms.ReturnInteger)
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

Private Sub txtWeb_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtEmail.SetFocus
        Case 40: KeyCode = 0: optTerms(0).SetFocus
    End Select
End Sub

Private Sub txtWeb_KeyPress(KeyAscii As MSForms.ReturnInteger)
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
