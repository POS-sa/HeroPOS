VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClaims 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   9630
   ClientLeft      =   -30
   ClientTop       =   810
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleMode       =   0  'User
   ScaleWidth      =   46257.31
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   12480
      TabIndex        =   25
      Top             =   7500
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   12480
      TabIndex        =   24
      Top             =   7890
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   12480
      TabIndex        =   23
      Top             =   8280
      Width           =   2175
   End
   Begin VSFlex8Ctl.VSFlexGrid grdClaims 
      Height          =   4350
      Left            =   90
      TabIndex        =   0
      Top             =   2790
      Width           =   15045
      _cx             =   26538
      _cy             =   7673
      Appearance      =   0
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
      BackColorSel    =   16642749
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmClaims.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   5670
         Width           =   1005
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   9330
      TabIndex        =   11
      Top             =   150
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20381697
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   9330
      TabIndex        =   12
      Top             =   480
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20381697
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   315
      Left            =   9330
      TabIndex        =   13
      Top             =   810
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20381697
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   315
      Left            =   9330
      TabIndex        =   14
      Top             =   1140
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20381697
      CurrentDate     =   38862
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11280
      TabIndex        =   28
      Top             =   7620
      Width           =   1005
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VAT:"
      Height          =   195
      Left            =   11280
      TabIndex        =   27
      Top             =   8010
      Width           =   1005
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   11280
      TabIndex        =   26
      Top             =   8400
      Width           =   1005
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11340
      TabIndex        =   22
      Top             =   270
      Width           =   1545
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VAT Number:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11340
      TabIndex        =   21
      Top             =   630
      Width           =   1545
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   315
      Left            =   12930
      TabIndex        =   20
      Top             =   150
      Width           =   2175
      VariousPropertyBits=   746604563
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "3836;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   315
      Left            =   12930
      TabIndex        =   19
      Top             =   510
      Width           =   2175
      VariousPropertyBits=   746604563
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "3836;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8070
      TabIndex        =   18
      Top             =   210
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8070
      TabIndex        =   17
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8070
      TabIndex        =   16
      Top             =   870
      Width           =   1005
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8070
      TabIndex        =   15
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1350
      TabIndex        =   10
      Top             =   810
      Width           =   2115
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. Number:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   900
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Number:"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   570
      Width           =   1035
   End
   Begin MSForms.ComboBox ComboBox2 
      Height          =   285
      Left            =   1350
      TabIndex        =   6
      Tag             =   "Up"
      Top             =   480
      Width           =   4065
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "7170;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1350
      TabIndex        =   5
      Top             =   150
      Width           =   2085
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1260
      Width           =   1035
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   795
      Left            =   1350
      TabIndex        =   3
      Top             =   1140
      Width           =   4065
      VariousPropertyBits=   746604563
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "7170;1402"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image2 
      Height          =   1935
      Left            =   7920
      Top             =   60
      Width           =   7245
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "12779;3413"
   End
   Begin MSForms.Image Image1 
      Height          =   585
      Index           =   2
      Left            =   60
      Top             =   2100
      Width           =   15105
      BackColor       =   16707305
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "26644;1032"
   End
   Begin MSForms.Image Image5 
      Height          =   2175
      Left            =   11160
      Top             =   7380
      Width           =   4005
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "7064;3836"
   End
   Begin MSForms.Image Image4 
      Height          =   2175
      Left            =   6750
      Top             =   7380
      Width           =   4305
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "7594;3836"
   End
   Begin MSForms.Image Image1 
      Height          =   2175
      Index           =   1
      Left            =   60
      Top             =   7380
      Width           =   6555
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "11562;3836"
   End
   Begin MSForms.Image Image3 
      Height          =   4515
      Left            =   60
      Top             =   2760
      Width           =   15105
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "26644;7964"
   End
   Begin MSForms.Image Image1 
      Height          =   1965
      Index           =   0
      Left            =   150
      Top             =   60
      Width           =   7695
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "13573;3466"
   End
End
Attribute VB_Name = "frmClaims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    grdClaims.RowHeight(0) = grdGRV.RowHeight(0) * 2
    grdClaims.Rows = 25
    grdClaims.Cols = 11
    grdClaims.TextMatrix(0, 0) = "Product Code"
    grdClaims.TextMatrix(0, 1) = "Description"
    grdClaims.TextMatrix(0, 2) = "Unit Size "
    grdClaims.TextMatrix(0, 3) = "Pack Size "
    grdClaims.TextMatrix(0, 4) = "Qty Ordered"
    grdClaims.TextMatrix(0, 5) = "Qty Deliverd"
    grdClaims.TextMatrix(0, 6) = "Qty Invoiced"
    grdClaims.TextMatrix(0, 7) = "Unit Price "
    grdClaims.TextMatrix(0, 8) = "Vat "
    grdClaims.TextMatrix(0, 9) = "Line Total "
    grdClaims.TextMatrix(0, 10) = "Claim Type"
    grdClaims.ColAlignment(2) = flexAlignRightCenter
    grdClaims.ColAlignment(3) = flexAlignRightCenter
    grdClaims.ColAlignment(4) = flexAlignRightCenter
    grdClaims.ColAlignment(5) = flexAlignRightCenter
    grdClaims.ColAlignment(6) = flexAlignRightCenter
    grdClaims.ColAlignment(7) = flexAlignRightCenter
    grdClaims.ColAlignment(8) = flexAlignRightCenter
    grdClaims.ColAlignment(9) = flexAlignRightCenter
    grdClaims.ColWidth(0) = 1200
    grdClaims.ColWidth(1) = 4000
    grdClaims.ColWidth(2) = 1150
    grdClaims.ColWidth(3) = 1150
    grdClaims.ColWidth(4) = 800
    grdClaims.ColWidth(5) = 800
    grdClaims.ColWidth(6) = 800
    grdClaims.ColWidth(7) = 1150
    grdClaims.ColWidth(8) = 1150
    grdClaims.ColWidth(9) = 800
End Sub




