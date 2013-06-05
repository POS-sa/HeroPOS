VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmOrder 
   Caption         =   "Purchase Order"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   13665
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10230
      Top             =   120
   End
   Begin VSFlex8Ctl.VSFlexGrid grdSupp 
      Height          =   150
      Left            =   60
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   13560
      _cx             =   23918
      _cy             =   265
      Appearance      =   2
      BorderStyle     =   1
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
      FormatString    =   $"frmOrder.frx":0000
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   11310
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   570
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65798147
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   11310
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65798147
      CurrentDate     =   38862
   End
   Begin VSFlex8Ctl.VSFlexGrid grdOrder 
      Height          =   4830
      Left            =   60
      TabIndex        =   1
      Top             =   2280
      Width           =   13565
      _cx             =   23927
      _cy             =   8520
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
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
      FormatString    =   $"frmOrder.frx":0078
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
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   5670
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx cmdSupplier 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   " Click to Search.... "
      Top             =   510
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Supplier..."
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
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   765
      Left            =   1590
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   900
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1349
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmOrder.frx":00F0
   End
   Begin btButtonEx.ButtonEx cmdStrip 
      Height          =   300
      Left            =   11670
      TabIndex        =   18
      Top             =   1890
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   529
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   4210752
      Caption         =   "Load Supplier Products"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VB.TextBox txtCalc3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox txtCalc2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtCalc1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   7680
      Width           =   1935
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   300
      Left            =   90
      TabIndex        =   43
      Top             =   1890
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   529
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   4210752
      Caption         =   "Delete Zero Quantities"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   6810
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   7650
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65798147
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker dtStop 
      Height          =   315
      Left            =   6810
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   8010
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   65798147
      CurrentDate     =   38862
   End
   Begin btButtonEx.ButtonEx cmdPDT 
      Height          =   300
      Left            =   10290
      TabIndex        =   53
      Top             =   1890
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   4210752
      Caption         =   "Load from PDT"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdPurchase 
      Height          =   285
      Left            =   3660
      TabIndex        =   54
      ToolTipText     =   " Click to Search.... "
      Top             =   7215
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      Appearance      =   3
      BorderColor     =   8421504
      Caption         =   "Purchase History..."
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
   Begin VB.TextBox txtShort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txtOver 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox txtReturn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1785
   End
   Begin btButtonEx.ButtonEx cmdGoodsReturned 
      Height          =   315
      Left            =   2370
      TabIndex        =   17
      ToolTipText     =   " Click to Search.... "
      Top             =   8370
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Goods Returned Claim..."
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
   Begin btButtonEx.ButtonEx cmdOverCharged 
      Height          =   315
      Left            =   2370
      TabIndex        =   16
      ToolTipText     =   " Click to Search.... "
      Top             =   8010
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Over Charged Claim..."
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
   Begin btButtonEx.ButtonEx cmdShortDelivered 
      Height          =   315
      Left            =   2400
      TabIndex        =   15
      ToolTipText     =   " Click to Search.... "
      Top             =   7650
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      Appearance      =   3
      Caption         =   "Short Delivered Claim..."
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
      Height          =   285
      Left            =   180
      TabIndex        =   55
      ToolTipText     =   " Click to Search.... "
      Top             =   7710
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Backorders..."
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
   Begin MSForms.Image Image3 
      Height          =   1215
      Left            =   60
      Top             =   7560
      Width           =   5655
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "9975;2143"
   End
   Begin MSForms.OptionButton optDate 
      Height          =   585
      Index           =   1
      Left            =   8730
      TabIndex        =   52
      Top             =   8160
      Width           =   975
      BackColor       =   16777215
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "1720;1032"
      Value           =   "0"
      Caption         =   "Last Month"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton optDate 
      Height          =   585
      Index           =   0
      Left            =   8730
      TabIndex        =   51
      Top             =   7590
      Width           =   975
      BackColor       =   16777215
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "1720;1032"
      Value           =   "1"
      Caption         =   "Last Week"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Consumtion:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5190
      TabIndex        =   49
      Top             =   8430
      Width           =   1545
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5730
      TabIndex        =   47
      Top             =   8070
      Width           =   1005
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5730
      TabIndex        =   46
      Top             =   7725
      Width           =   1005
   End
   Begin VB.Label lblFax 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   11460
      TabIndex        =   42
      Top             =   1410
      Width           =   2025
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9720
      TabIndex        =   41
      Top             =   1410
      Width           =   1545
   End
   Begin VB.Label lblTel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7950
      TabIndex        =   40
      Top             =   1410
      Width           =   2025
   End
   Begin MSForms.Image Image14 
      Height          =   315
      Left            =   11340
      Top             =   8010
      Width           =   2205
      BorderColor     =   12632256
      Size            =   "3889;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image15 
      Height          =   315
      Left            =   11340
      Top             =   8370
      Width           =   2205
      BorderColor     =   12632256
      Size            =   "3889;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image13 
      Height          =   315
      Left            =   11340
      Top             =   7650
      Width           =   2205
      BorderColor     =   12632256
      Size            =   "3889;556"
      VariousPropertyBits=   19
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Number:"
      Height          =   195
      Left            =   150
      TabIndex        =   39
      Top             =   240
      Width           =   1215
   End
   Begin MSForms.ComboBox cmbSuppliers 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Tag             =   "Up"
      Top             =   510
      Width           =   4515
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "7964;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      BorderColor     =   12632256
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   165
      Left            =   390
      TabIndex        =   38
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Order Number:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6210
      TabIndex        =   37
      Top             =   240
      Width           =   1545
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone Number:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6210
      TabIndex        =   36
      Top             =   1410
      Width           =   1545
   End
   Begin VB.Label lblGRV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      TabIndex        =   35
      Top             =   230
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10260
      TabIndex        =   34
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
      Left            =   10260
      TabIndex        =   33
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10260
      TabIndex        =   32
      Top             =   1020
      Width           =   1005
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vat Number:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6210
      TabIndex        =   31
      Top             =   630
      Width           =   1545
   End
   Begin VB.Label lblOrder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7950
      TabIndex        =   30
      Top             =   270
      Width           =   1545
   End
   Begin VB.Label lblVat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7950
      TabIndex        =   29
      Top             =   660
      Width           =   1605
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
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6210
      TabIndex        =   28
      Top             =   1020
      Width           =   1545
   End
   Begin VB.Label lblContact 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7950
      TabIndex        =   27
      Top             =   1020
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   10275
      TabIndex        =   26
      Top             =   8430
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vat:"
      Height          =   195
      Left            =   10275
      TabIndex        =   25
      Top             =   8055
      Width           =   1005
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10275
      TabIndex        =   24
      Top             =   7680
      Width           =   1005
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
      Left            =   9900
      TabIndex        =   23
      Top             =   7245
      Width           =   3675
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding Claims"
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
      Left            =   210
      TabIndex        =   22
      Top             =   7230
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.Label DTPicker1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11400
      TabIndex        =   20
      Top             =   270
      Width           =   1995
   End
   Begin VB.Label lblAve 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6750
      TabIndex        =   19
      Top             =   1920
      Width           =   3465
   End
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   11310
      Top             =   180
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3942;609"
   End
   Begin MSForms.Image Image9 
      Height          =   345
      Left            =   7800
      Top             =   570
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3942;609"
   End
   Begin MSForms.Image Image8 
      Height          =   345
      Left            =   7800
      Top             =   180
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3942;609"
   End
   Begin MSForms.Image Image11 
      Height          =   345
      Left            =   7800
      Top             =   960
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3942;609"
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
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      Top             =   1920
      Width           =   4785
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   60
      Top             =   1860
      Width           =   13575
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "23945;661"
   End
   Begin MSForms.Image Image5 
      Height          =   1215
      Index           =   0
      Left            =   9840
      Top             =   7560
      Width           =   3795
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6694;2143"
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
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   9840
      Top             =   7170
      Width           =   3795
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6694;661"
   End
   Begin MSForms.Image Image19 
      Height          =   345
      Left            =   7800
      Top             =   1350
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3942;609"
   End
   Begin MSForms.Image Image21 
      Height          =   345
      Left            =   11310
      Top             =   1350
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3942;609"
   End
   Begin MSForms.Image Image2 
      Height          =   1755
      Left            =   6090
      Top             =   60
      Width           =   7545
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "13309;3096"
   End
   Begin VB.Label lblCon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consumption"
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
      Left            =   5820
      TabIndex        =   50
      Top             =   7245
      Width           =   3915
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   4
      Left            =   5760
      Top             =   7170
      Width           =   4035
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7117;661"
   End
   Begin VB.Label lblConsump 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6930
      TabIndex        =   48
      Top             =   8430
      Width           =   1545
   End
   Begin MSForms.Image Image4 
      Height          =   315
      Left            =   6810
      Top             =   8370
      Width           =   1755
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3096;556"
   End
   Begin MSForms.Image Image5 
      Height          =   1215
      Index           =   1
      Left            =   5760
      Top             =   7560
      Width           =   4035
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7117;2143"
   End
   Begin MSForms.Image Image6 
      Height          =   8835
      Left            =   0
      Top             =   0
      Width           =   13710
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "24183;15584"
   End
   Begin MSForms.Image Image16 
      Height          =   315
      Left            =   1980
      Top             =   7650
      Visible         =   0   'False
      Width           =   2055
      BorderColor     =   12632256
      Size            =   "3625;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image17 
      Height          =   315
      Left            =   1980
      Top             =   8370
      Visible         =   0   'False
      Width           =   2055
      BorderColor     =   12632256
      Size            =   "3625;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image18 
      Height          =   315
      Left            =   1980
      Top             =   8010
      Visible         =   0   'False
      Width           =   2055
      BorderColor     =   12632256
      Size            =   "3625;556"
      VariousPropertyBits=   19
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub OpenOrder()
    Load frmOpen
    frmOpen.Tag = "Order"
    DoEvents
    frmOpen.Show vbModal
End Sub
Public Sub Accept_Order()
    If cmbSuppliers.Text = "" Then
        MsgBox "Please Select a Supplier Number", vbCritical, "HeroPOS"
        cmbSuppliers.SetFocus
        Exit Sub
    End If
    frmMain.Toolbar1.Tag = ""
    Select Case MsgBox("Do you want to place the current order?", vbYesNoCancel, "HeroPOS")
        Case vbYes
            Screen.MousePointer = 11
            If Val(lblOrder.Caption) <> 0 Then
                ActiveUpdateServer "Delete from Purchase_Order_Listing where Order_No = " & Val(lblOrder.Caption)
            Else
                ActiveReadServer1 "Select isnull(max(Order_No),0) + 1 as Order_No from Purchase_Order_Journal"
                lblOrder.Caption = Format(rs1.Fields("Order_No"), "000000")
                rs1.Close
            End If
            DoEvents
            For i = 1 To grdOrder.Rows - 1
                If Trim(grdOrder.TextMatrix(i, 0)) <> "" Then
                    ActiveUpdateServer "INSERT INTO Purchase_Order_Journal (Supplier_No,Workstation_No, User_No, Location_No, Order_No, Product_Code, Department_No, Pack_Size, Qty_Ordered, Price_Ordered, Vat_Rate, Line_Total,Date_Time,Order_Date, Delivery_Date)" & _
                    " VALUES ('" & lblGRV.Caption & "'," & Workstation_No & "," & UserRecord.User_Number & ",1," & Val(lblOrder.Caption) & ",'" & grdOrder.TextMatrix(i, 0) & "','" & grdOrder.TextMatrix(i, 9) & "'," & grdOrder.TextMatrix(i, 3) & "," & grdOrder.TextMatrix(i, 5) & "," & grdOrder.TextMatrix(i, 6) & "," & Replace(grdOrder.TextMatrix(i, 7), "%", "") & "," & grdOrder.ValueMatrix(i, 8) & ",GetDate(),'" & DTPicker2.Value & "','" & DTPicker3.Value & "')"
                End If
            Next i
            MsgBox "Order no: " & lblOrder.Caption & " Placed Successfully", vbInformation, "HeroPOS"
            Screen.MousePointer = 0
            frmMain.Toolbar1.Tag = lblOrder.Caption
            Screen.MousePointer = 0
            Form_Load
        Case vbNo
            frmMain.picProdBar.Visible = False
            Unload frmOrder
            DoEvents
            frmMain.cmdBar(7).Enabled = True
            frmDetails.Show
        Case Else
    End Select
End Sub
Public Sub SaveOrder()
    If cmbSuppliers.Text = "" Then
        MsgBox "Please Enter a Supplier First", vbCritical, "HeroPOS"
        cmbSuppliers.SetFocus
        Exit Sub
    End If
    Select Case MsgBox("Do you want to Save the Current Order?", vbYesNoCancel, "HeroPOS")
        Case vbYes
            Screen.MousePointer = 11
            If Val(lblOrder.Caption) <> 0 Then
                ActiveUpdateServer "Delete from Purchase_Order_Listing where Order_No = " & Val(lblOrder.Caption)
            Else
                ActiveReadServer1 "Select isnull(max(Order_No),0) + 1 as Order_No from Purchase_Order_Journal"
                lblOrder.Caption = Format(rs1.Fields("Order_No"), "000000")
                rs1.Close
            End If
            ActiveUpdateServer "Insert into Purchase_Order_Journal (Order_No)  values (" & Val(lblOrder.Caption) & ")"
            DoEvents
            For i = 1 To grdOrder.Rows - 1
                ActiveUpdateServer "INSERT INTO Purchase_Order_Listing(Supplier_No,Workstation_No, User_No, Location_No, Order_No, Product_Code, Department_No, Pack_Size, Qty_Ordered, Price_Ordered, Vat_Rate, Line_Total,Date_Time,Order_Date, Delivery_Date)" & _
                " VALUES ('" & lblGRV.Caption & "'," & Workstation_No & "," & UserRecord.User_Number & ",1," & Val(lblOrder.Caption) & ",'" & grdOrder.TextMatrix(i, 0) & "','" & grdOrder.TextMatrix(i, 9) & "'," & grdOrder.TextMatrix(i, 3) & "," & grdOrder.TextMatrix(i, 5) & "," & grdOrder.TextMatrix(i, 6) & "," & Replace(grdOrder.TextMatrix(i, 7), "%", "") & "," & grdOrder.ValueMatrix(i, 8) & ",GetDate(),'" & DTPicker2.Value & "','" & DTPicker3.Value & "')"
            Next i
            MsgBox "Order no: " & lblOrder.Caption & " saved Successfully", vbInformation, "HeroPOS"
            Screen.MousePointer = 0
        Case vbNo
            frmMain.picProdBar.Visible = False
            Unload frmOrder
            DoEvents
            frmMain.cmdBar(7).Enabled = True
            frmDetails.Show
        Case Else
    End Select
End Sub
Public Sub DeleteOrder()
    Select Case MsgBox("Are you sure you want to Delete the Order", vbYesNo, "HeroPOS")
        Case vbYes
            ActiveUpdateServer "Delete from Purchase_Order_Listing where Order_No = " & Val(lblOrder.Caption)
            ActiveUpdateServer "Delete from Purchase_Order_Journal where Order_No = " & Val(lblOrder.Caption)
            frmMain.picProdBar.Visible = False
            Unload frmOrder
            DoEvents
            frmMain.cmdBar(7).Enabled = True
            frmDetails.Show
        Case Else
    End Select
End Sub

Private Sub ButtonEx1_Click()
top:
    For i = 1 To grdOrder.Rows - 1
        If grdOrder.ValueMatrix(i, 5) = 0 Then
            grdOrder.RemoveItem (i)
            GoTo top
        End If
    Next i
End Sub


Private Sub ButtonEx2_Click()
frmOrder.Hide
ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Back Orders')"
'frmBackorders.Show
End Sub

Private Sub cmbSuppliers_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbSuppliers_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            Else
                grdOrder_Click
            End If
        Case 38
            DoEvents
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                If grdOrder.Rows = 1 Then
                    cmbSuppliers.SetFocus
                Else
                    grdOrder.SetFocus
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
End Sub
Private Sub cmbSuppliers_LostFocus()
    If cmbSuppliers.Text <> "" Then
        txtAddress.Tag = Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "")
        LoadInfo
        grdOrder_Click
    End If
End Sub
Private Sub cmdPdt_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    stime = 0
    frmOrder.Tag = "1"
    Kill App.Path & "\Cipher\Test.txt"
    DoEvents
    Shell App.Path & "\Cipher\Data_Read.exe", vbNormalFocus
    DoEvents
    stime = Timer
    While Timer - stime < 12: Wend
    On Error GoTo trap
    filenum = FreeFile
    Open App.Path & "\Cipher\Test.txt" For Input As filenum
    i = 0
    Barcode = ""
    While Not EOF(filenum)
        Line Input #filenum, newline
        Barcode = Trim(Mid(newline, 1, InStr(newline, ",") - 1))
        Qty = Val(Trim(Mid(newline, InStr(newline, ",") + 1)))
        ActiveReadServer "Select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
        "+ Unit_of_Measure END AS Description,Stock_Item,Pack_Size,Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & Barcode & "'"
        If rs.RecordCount > 0 Then
            If grdOrder.TextMatrix(grdOrder.Row, 2) <> "" Then
                grdOrder.Rows = grdOrder.Rows + 1
            End If
            grdOrder.Row = grdOrder.Rows - 1
            grdOrder.TextMatrix(grdOrder.Row, 0) = Barcode
            grdOrder.TextMatrix(grdOrder.Row, 3) = "1"
            grdOrder.TextMatrix(grdOrder.Row, 4) = "0"
            grdOrder.TextMatrix(grdOrder.Row, 5) = Qty
            grdOrder.TextMatrix(grdOrder.Row, 2) = rs.Fields("Description") & " - " & Barcode
            grdOrder.TextMatrix(grdOrder.Row, 6) = rs.Fields("Landed_Cost")
            grdOrder.TextMatrix(grdOrder.Row, 7) = rs.Fields("Sales_Tax") & "%"
            grdOrder.TextMatrix(grdOrder.Row, 9) = rs.Fields("Department_No")
            grdOrder.TextMatrix(grdOrder.Row, 3) = rs.Fields("Pack_Size")
            grdOrder.TextMatrix(grdOrder.Row, 8) = grdOrder.ValueMatrix(grdOrder.Row, 5) * grdOrder.ValueMatrix(grdOrder.Row, 6)
            If rs.Fields("Pack_Size") = 1 And rs.Fields("Stock_Item") = 1 Then
                ActiveReadServer2 "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & Barcode & "'"
                If rs2.RecordCount > 0 Then SOH = Val(rs2.Fields("SOH") & "")
            Else
                ActiveReadServer2 "Select Link_Code,(Select Pack_Size from Products where Pack_Links.Product_Code=Products.Product_Code) as Pack_Size from Pack_Links where Product_Code = '" & Barcode & "'"
                If rs2.RecordCount > 0 Then
                    ActiveReadServer2 "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                    "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & rs2.Fields("Link_Code") & "'"
                    SOH = Val(rs2.Fields("SOH") & "")
                End If
                rs2.Close
                ActiveReadServer2 "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & Barcode & "'"
            End If
            If rs2.RecordCount > 0 Then
                AveCost = Val(rs2.Fields("Ave_Cost") & "")
                lblInfo.Caption = "Last Landed Cost: " & Format(rs2.Fields("Landed_Cost"), "0.00") & "   Ave Cost: " & Format(rs2.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs2.Fields("Selling_Price"), "0.00")
            End If
            rs2.Close
            
            NEWQTY = grdOrder.ValueMatrix(grdOrder.Row, 5)
            Land = grdOrder.ValueMatrix(grdOrder.Row, 6)
            If SOH < 0 Then SOH = 0
               
            If NEWQTY <> 0 And Land <> 0 Then
                NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                    lblAve.ForeColor = &HC0&
                    lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                Else
                    lblAve.ForeColor = &HC00000
                    lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                End If
            Else
                lblAve.Caption = "New Ave Cost = " & Format(Land, "0.00")
            End If
            grdOrder.TextMatrix(grdOrder.Row, 4) = SOH
            GetConsumption
        End If
        rs.Close
    Wend
    Close filenum
    grdOrder.Col = 5
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(3).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = True
    Calculate_Totals
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
trap:
    Close filenum
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub cmdPurchase_Click()
    If Trim(grdOrder.TextMatrix(grdOrder.Row, 0)) = "" Then Exit Sub
    frmOrder.Tag = "1"
    Load frmPurchase
    frmPurchase.Tag = "Order"
    DoEvents
    frmPurchase.Show vbModal
End Sub

Private Sub cmdStrip_Click()
    Screen.MousePointer = 11
    grdOrder.Rows = 1
    ActiveReadServer3 "Select * from Suppliers_Links_View where Supplier_No = '" & lblGRV.Caption & "' order by Description"
    While Not rs3.EOF
        grdOrder.Rows = grdOrder.Rows + 1
        grdOrder.TextMatrix(grdOrder.Rows - 1, 0) = rs3.Fields("Product_Code")
        grdOrder.TextMatrix(grdOrder.Rows - 1, 1) = rs3.Fields("Supplier_Code")
        grdOrder.TextMatrix(grdOrder.Rows - 1, 2) = rs3.Fields("Description") & " - " & rs3.Fields("Product_Code")
        grdOrder.TextMatrix(grdOrder.Rows - 1, 3) = "1"
        grdOrder.TextMatrix(grdOrder.Rows - 1, 5) = "0"
        grdOrder.TextMatrix(grdOrder.Rows - 1, 6) = rs3.Fields("Landed_Cost")
        rs3.MoveNext
    Wend
    rs3.Close
    For i = 1 To grdOrder.Rows - 1
        ActiveReadServer "Select (Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH,Pack_Size,Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdOrder.TextMatrix(i, 0) & "'"
        If rs.RecordCount > 0 Then
            grdOrder.TextMatrix(i, 4) = Round(Val(rs.Fields("SOH") & ""), 3)
            grdOrder.TextMatrix(i, 6) = rs.Fields("Landed_Cost") * rs.Fields("Pack_Size")
            grdOrder.TextMatrix(i, 7) = rs.Fields("Sales_Tax") & "%"
            grdOrder.TextMatrix(i, 9) = rs.Fields("Department_No")
            grdOrder.TextMatrix(i, 8) = "0.00"
            grdOrder.TextMatrix(i, 3) = rs.Fields("Pack_Size")
        End If
        rs.Close
    Next i
    If grdOrder.Rows > 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
    End If
    grdOrder.SetFocus
    Screen.MousePointer = 0
End Sub
Private Sub cmdSupplier_Click()
    Select Case cmdSupplier.Value
        Case 0
            Screen.MousePointer = 11
            grdOrder.Rows = 1
            grdSupp.Visible = True
            LoadSuppliers
            DTPicker1.Enabled = False
            DTPicker2.Enabled = False
            DTPicker3.Enabled = False
            cmbSuppliers.Enabled = False
            DoEvents
            If grdSupp.Rows > 1 Then
                grdSupp.Col = 1
                grdSupp.SetFocus
            End If
            Screen.MousePointer = 0
        Case 1
            grdSupp.Visible = False
            cmbSuppliers.Enabled = True
            DTPicker1.Enabled = True
            DTPicker2.Enabled = True
            DTPicker3.Enabled = True
            DoEvents
    End Select
End Sub
Public Sub LoadOrder(Order_No)
    ActiveReadServer "Select * from Purchase_Order_Listing where Order_No = " & Order_No
    If rs.RecordCount > 0 Then
        lblOrder.Caption = Format(Order_No, "000000")
        txtAddress.Tag = rs.Fields("Supplier_No")
        lblOrder.Caption = Format(rs.Fields("Order_No"), "000000")
        lblGRV.Caption = rs.Fields("Supplier_No")
        DTPicker1.Caption = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
        DTPicker2.Value = rs.Fields("Order_Date")
        DTPicker3.Value = rs.Fields("Delivery_Date")
        ActiveReadServer1 "Select * from Suppliers where Supplier_No='" & txtAddress.Tag & "'"
        If rs1.RecordCount > 0 Then
            cmbSuppliers.Text = rs1.Fields("Supplier_Name") & " (" & rs1.Fields("Supplier_No") & ")"
            txtAddress.Text = rs1.Fields("Address")
            lblContact.Caption = rs1.Fields("Contact_Person")
            lblVat.Caption = rs1.Fields("VAT_No") & ""
            lblTel.Caption = rs1.Fields("Business_Tel") & ""
            lblFax.Caption = rs1.Fields("Fax_Tel") & ""
        End If
        rs1.Close
    End If
    i = 0
    grdOrder.Rows = 1
    While Not rs.EOF
        i = i + 1
        grdOrder.Rows = grdOrder.Rows + 1
        grdOrder.TextMatrix(i, 0) = rs.Fields("Product_Code")
        ActiveReadServer1 "Select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
        "+ Unit_of_Measure END AS Description from Products where Product_Code='" & rs.Fields("Product_Code") & "'"
        If rs1.RecordCount > 0 Then
            grdOrder.TextMatrix(i, 2) = rs1.Fields("Description") & " - " & rs.Fields("Product_Code")
        Else
            grdOrder.TextMatrix(i, 2) = "Unknown Product - " & rs.Fields("Product_Code")
        End If
        rs1.Close
        grdOrder.TextMatrix(i, 3) = rs.Fields("Pack_Size")
        grdOrder.TextMatrix(i, 5) = Round(rs.Fields("Qty_Ordered"), 3)
        grdOrder.TextMatrix(i, 6) = Format(rs.Fields("Price_Ordered"), "0.00")
        grdOrder.TextMatrix(i, 7) = rs.Fields("Vat_Rate") & "%"
        grdOrder.TextMatrix(i, 9) = rs.Fields("Department_No")
        grdOrder.TextMatrix(i, 8) = Format(rs.Fields("Line_Total"), "0.00")
        ActiveReadServer2 "Select Supplier_Code from Suppliers_Links_View where Supplier_No = '" & lblGRV.Caption & "' and Product_Code = '" & rs.Fields("Product_Code") & "' order by Description"
        If rs2.RecordCount > 0 Then
            grdOrder.TextMatrix(i, 1) = rs2.Fields("Supplier_Code")
        End If
        rs2.Close
        ActiveReadServer2 "Select sum(isnull(Stock_on_Hand,0)) as SOH from Quantities where Product_Code = '" & grdOrder.TextMatrix(i, 0) & "'"
        If rs2.RecordCount > 0 Then
            grdOrder.TextMatrix(i, 4) = Round(Val(rs2.Fields("SOH") & ""), 3)
        End If
        rs2.Close
        rs.MoveNext
    Wend
    rs.Close
    Calculate_Totals
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(3).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = True
End Sub

Private Sub dtStart_Change()
    optDate(0).Value = 0
    optDate(1).Value = 0
    GetConsumption
End Sub
Private Sub dtStop_Change()
    optDate(0).Value = 0
    optDate(1).Value = 0
    GetConsumption
End Sub
Private Sub Form_Activate()
    If frmOrder.Tag <> "" Then Exit Sub
    frmMain.Toolbar1.Buttons(11).Enabled = False
    frmMain.Toolbar1.Buttons(12).Enabled = False
    DTStart.Value = DateAdd("d", -7, Date)
    DTStop.Value = Date
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Unload rptOrder
    lblAve.Caption = ""
    DTPicker1.Caption = Format(Date, "ddd dd MMM yyyy")
    DTPicker2.Value = Date
    DTPicker3.Value = Date
    DTPicker4.Value = Date
    grdSupp.Height = 4830
    grdOrder.RowHeight(0) = 550
    grdOrder.Rows = 1
    grdOrder.Cols = 10
    grdOrder.ColHidden(0) = True
    grdOrder.ColHidden(9) = True
    grdOrder.TextMatrix(0, 0) = "Product Code"
    grdOrder.TextMatrix(0, 1) = "Supplier Code"
    grdOrder.TextMatrix(0, 2) = "Description"
    grdOrder.TextMatrix(0, 3) = "Pack Size "
    grdOrder.TextMatrix(0, 4) = "Stock on Hand"
    grdOrder.TextMatrix(0, 5) = "Order Qty"
    grdOrder.TextMatrix(0, 6) = "Unit Price "
    grdOrder.TextMatrix(0, 7) = "Vat "
    grdOrder.TextMatrix(0, 8) = "Line Total "
    grdOrder.ColAlignment(1) = flexAlignLeftCenter
    grdOrder.ColAlignment(2) = flexAlignLeftCenter
    grdOrder.ColAlignment(3) = flexAlignRightCenter
    grdOrder.ColAlignment(4) = flexAlignRightCenter
    grdOrder.ColAlignment(5) = flexAlignRightCenter
    grdOrder.ColAlignment(6) = flexAlignRightCenter
    grdOrder.ColAlignment(7) = flexAlignRightCenter
    grdOrder.ColAlignment(8) = flexAlignRightCenter
    grdOrder.ColWidth(1) = grdOrder.Width * 0.13
    grdOrder.ColWidth(2) = grdOrder.Width * 0.3
    grdOrder.ColWidth(3) = grdOrder.Width * 0.09
    grdOrder.ColWidth(4) = grdOrder.Width * 0.09
    grdOrder.ColWidth(5) = grdOrder.Width * 0.09
    grdOrder.ColWidth(6) = grdOrder.Width * 0.09
    grdOrder.ColWidth(7) = grdOrder.Width * 0.09
    grdOrder.ColWidth(8) = grdOrder.Width * 0.11
    grdOrder.ColFormat(6) = "0.00"
    grdOrder.ColFormat(8) = "0.00"
    cmbSuppliers.Clear
    ActiveReadServer "Select Supplier_No, Supplier_Name from Suppliers order by Supplier_Name"
    While Not rs.EOF
        cmbSuppliers.AddItem rs.Fields("Supplier_Name") & " (" & rs.Fields("Supplier_No") & ")"
        rs.MoveNext
    Wend
    rs.Close
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
    grdSupp.ColAlignment(0) = flexAlignLeftCenter
    grdSupp.ColAlignment(1) = flexAlignLeftCenter
    grdSupp.ColAlignment(2) = flexAlignLeftCenter
    grdSupp.ColAlignment(3) = flexAlignLeftCenter
    grdSupp.ColAlignment(4) = flexAlignRightCenter
    lblOrder.Caption = "000000"
    lblVat.Caption = ""
    lblContact.Caption = ""
    lblOrder.Caption = "000000"
    txtAddress.Text = ""
    txtInvoice.Text = ""
    txtSubtotal.Text = "0.00"
    txtVat.Text = "0.00"
    txtTotal.Text = "0.00"
    txtCalc1.Text = "0.00"
    txtCalc2.Text = "0.00"
    txtCalc3.Text = "0.00"
    txtShort.Text = "0.00"
    txtOver.Text = "0.00"
    txtReturn.Text = "0.00"
    chkShort.Value = 0
    chkOver.Value = 0
    chkReturn.Value = 0
    frmMain.Toolbar1.Buttons(2).Caption = "Place"
    frmMain.Toolbar1.Buttons(2).Enabled = False
    frmMain.Toolbar1.Buttons(3).Enabled = True
    frmMain.Toolbar1.Buttons(4).Enabled = False
    txtDiscount.Text = "0.00"
    txtUllages.Text = "0.00"
    txtTransport.Text = "0.00"
    picSundries.Visible = False
    Label18 = "Calculated Totals"
    On Error GoTo 0
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(2).Caption = "New"
    Unload frmSearch
End Sub

Private Sub grdOrder_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Select Case Col
        Case 1
            grdOrder.TextMatrix(grdOrder.Row, 3) = "1"
            grdOrder.TextMatrix(grdOrder.Row, 4) = "0"
            grdOrder.TextMatrix(grdOrder.Row, 5) = "0"
            ActiveReadServer "Select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
            "+ Unit_of_Measure END AS Description,Pack_Size,Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
            If rs.RecordCount > 0 Then
                grdOrder.TextMatrix(grdOrder.Row, 2) = rs.Fields("Description") & " - " & grdOrder.TextMatrix(grdOrder.Row, 0)
                grdOrder.TextMatrix(grdOrder.Row, 6) = rs.Fields("Landed_Cost") * rs.Fields("Pack_Size")
                grdOrder.TextMatrix(grdOrder.Row, 7) = rs.Fields("Sales_Tax") & "%"
                grdOrder.TextMatrix(grdOrder.Row, 9) = rs.Fields("Department_No")
                grdOrder.TextMatrix(grdOrder.Row, 3) = rs.Fields("Pack_Size")
            End If
            rs.Close
            grdOrder.Col = 5
            frmMain.Toolbar1.Buttons(2).Enabled = True
            frmMain.Toolbar1.Buttons(3).Enabled = False
            frmMain.Toolbar1.Buttons(4).Enabled = True
        Case 4
            If grdOrder.ValueMatrix(grdOrder.Row, 0) <> 0 Then
                ActiveReadServer "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    AveCost = Val(rs.Fields("Ave_Cost") & "")
                    SOH = Val(rs.Fields("SOH") & "")
                    lblInfo.Caption = "Last Landed Cost: " & Format(rs.Fields("Landed_Cost"), "0.00") & "   Ave Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs.Fields("Selling_Price"), "0.00")
                End If
                rs.Close
                
                NEWQTY = grdOrder.ValueMatrix(grdOrder.Row, 6)
                Land = grdOrder.ValueMatrix(grdOrder.Row, 7)
                If SOH < 0 Then SOH = 0
                   
                If NEWQTY <> 0 And Land <> 0 Then
                    NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                    If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                        lblAve.ForeColor = &HC0&
                        lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    Else
                        lblAve.ForeColor = &HC00000
                        lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    End If
                Else
                    lblAve.Caption = "New Ave Cost = " & Format(Land, "0.00")
                End If
            End If
        Case 5
            grdOrder.TextMatrix(grdOrder.Row, 8) = grdOrder.ValueMatrix(grdOrder.Row, 5) * grdOrder.ValueMatrix(grdOrder.Row, 6)
        Case 6
            If grdOrder.ValueMatrix(grdOrder.Row, 0) <> 0 Then
                grdOrder.TextMatrix(grdOrder.Row, 8) = grdOrder.ValueMatrix(grdOrder.Row, 5) * grdOrder.ValueMatrix(grdOrder.Row, 6)
                ActiveReadServer "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    AveCost = Val(rs.Fields("Ave_Cost") & "")
                    SOH = Val(rs.Fields("SOH") & "")
                    lblInfo.Caption = "Last Landed Cost: " & Format(rs.Fields("Landed_Cost"), "0.00") & "   Ave Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs.Fields("Selling_Price"), "0.00")
                End If
                rs.Close
                
                NEWQTY = grdOrder.ValueMatrix(grdOrder.Row, 5)
                Land = grdOrder.ValueMatrix(grdOrder.Row, 6)
                If SOH < 0 Then SOH = 0
                   
                If NEWQTY <> 0 And Land <> 0 Then
                    NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                    If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                        lblAve.ForeColor = &HC0&
                        lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    Else
                        lblAve.ForeColor = &HC00000
                        lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    End If
                Else
                    lblAve.Caption = "New Ave Cost = " & Format(Land, "0.00")
                End If
            End If
            If grdOrder.ValueMatrix(grdOrder.Row, 5) <> 0 Then
                grdOrder.TextMatrix(grdOrder.Row, 8) = grdOrder.ValueMatrix(grdOrder.Row, 5) * grdOrder.ValueMatrix(grdOrder.Row, 6)
            End If
        Case 8
            If grdOrder.ValueMatrix(grdOrder.Row, 5) <> 0 Then
                grdOrder.TextMatrix(grdOrder.Row, 6) = grdOrder.ValueMatrix(grdOrder.Row, 8) / grdOrder.ValueMatrix(grdOrder.Row, 5)
            End If
            If grdOrder.ValueMatrix(grdOrder.Row, 0) <> 0 Then
                grdOrder.TextMatrix(grdOrder.Row, 8) = grdOrder.ValueMatrix(grdOrder.Row, 5) * grdOrder.ValueMatrix(grdOrder.Row, 6)
                ActiveReadServer "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    AveCost = Val(rs.Fields("Ave_Cost") & "")
                    SOH = Val(rs.Fields("SOH") & "")
                    lblInfo.Caption = "Last Landed Cost: " & Format(rs.Fields("Landed_Cost"), "0.00") & "   Ave Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs.Fields("Selling_Price"), "0.00")
                End If
                rs.Close
                
                NEWQTY = grdOrder.ValueMatrix(grdOrder.Row, 5)
                Land = grdOrder.ValueMatrix(grdOrder.Row, 6)
                If SOH < 0 Then SOH = 0
                If NEWQTY <> 0 And Land <> 0 Then
                    NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                    If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                        lblAve.ForeColor = &HC0&
                        lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    Else
                        lblAve.ForeColor = &HC00000
                        lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    End If
                Else
                    lblAve.Caption = "New Ave Cost = " & Format(Land, "0.00")
                End If
            End If
    End Select
    On Error GoTo 0
    Calculate_Totals
End Sub
Private Sub Calculate_Totals()
    txtCalc1.Text = "0.00"
    txtCalc2.Text = "0.00"
    txtCalc3.Text = "0.00"
    For i = 1 To grdOrder.Rows - 1
        txtCalc1.Text = Format(Val(txtCalc1.Text) + grdOrder.ValueMatrix(i, 8), "0.00")
        txtCalc3.Text = Format(Val(txtCalc3.Text) + (grdOrder.ValueMatrix(i, 8) * (1 + grdOrder.ValueMatrix(i, 7))), "0.00")
    Next i
    txtCalc2.Text = Format(Val(txtCalc3.Text) - Val(txtCalc1.Text), "0.00")
End Sub
Private Sub grdOrder_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If cmbSuppliers.Text = "" Then
        cmdStrip.Enabled = False
        cmdPdt.Enabled = False
        ButtonEx1.Enabled = False
    Else
        cmdStrip.Enabled = True
        cmdPdt.Enabled = True
        ButtonEx1.Enabled = True
    End If
    If grdOrder.TextMatrix(NewRow, 0) <> "" And NewRow <> 0 Then
        If grdOrder.ValueMatrix(NewRow, 0) <> 0 Then
            If grdOrder.ValueMatrix(NewRow, 3) = 1 Then
                ActiveReadServer2 "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdOrder.TextMatrix(NewRow, 0) & "'"
                If rs2.RecordCount > 0 Then SOH = Val(rs2.Fields("SOH") & "")
            Else
                ActiveReadServer2 "Select Link_Code,(Select Pack_Size from Products where Pack_Links.Product_Code=Products.Product_Code) as Pack_Size from Pack_Links where Product_Code = '" & grdOrder.TextMatrix(NewRow, 0) & "'"
                If rs2.RecordCount > 0 Then
                    ActiveReadServer2 "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                    "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & rs2.Fields("Link_Code") & "'"
                    SOH = Val(rs2.Fields("SOH") & "")
                Else
                    rs2.Close
                    ActiveReadServer2 "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                    "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdOrder.TextMatrix(NewRow, 0) & "'"
                    If rs2.RecordCount > 0 Then SOH = Val(rs2.Fields("SOH") & "")
                End If
                rs2.Close
                ActiveReadServer2 "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdOrder.TextMatrix(NewRow, 0) & "'"
            End If
            If rs2.RecordCount > 0 Then
                AveCost = Val(rs2.Fields("Ave_Cost") & "")
                grdOrder.TextMatrix(NewRow, 4) = Round(SOH, 3)
                lblInfo.Caption = "Last Landed Cost: " & Format(rs2.Fields("Landed_Cost"), "0.00") & "   Ave Cost: " & Format(rs2.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs2.Fields("Selling_Price"), "0.00")
            End If
            rs2.Close
            
            NEWQTY = grdOrder.ValueMatrix(NewRow, 5)
            Land = grdOrder.ValueMatrix(NewRow, 6)
            If SOH < 0 Then SOH = 0
               
            If NEWQTY <> 0 And Land <> 0 Then
                NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                    lblAve.ForeColor = &HC0&
                    lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                Else
                    lblAve.ForeColor = &HC00000
                    lblAve.Caption = "New Ave Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                End If
            Else
                lblAve.Caption = "New Ave Cost = " & Format(Land, "0.00")
            End If
        End If
        GetConsumption
    End If
End Sub
Private Sub GetConsumption()
    On Error Resume Next
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, DTStop.Value)
    Else
        Selender = DTStop.Value
    End If
    ActiveReadServer2 "SELECT SUM(Sales_Journal.Qty) AS Qty FROM Sales_Journal WHERE (Sales_Journal.Function_Key IN (7)) AND ((ISNULL(Sales_Journal.Extra, '') = '' or Sales_Journal.Extra='Return Item')) and Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "' and " & _
    "(Date_Time > '" & DTStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    lblConsump.Caption = Round(Val(rs2.Fields("Qty") & ""), 3)
    rs2.Close
    DoEvents
    ActiveReadServer2 "SELECT SUM(Qty_Consumed) AS Qty FROM Consumption_Journal WHERE Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "' and " & _
    "(Date_Time > '" & DTStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    lblConsump.Caption = Round(Val(rs2.Fields("Qty") & "") + Val(lblConsump.Caption), 3)
    rs2.Close
    grdOrder.SetFocus
    On Error GoTo 0
End Sub
Private Sub grdOrder_Click()
    If cmbSuppliers.Text <> "" And Val(lblOrder.Caption) = 0 Then
        If grdOrder.Rows = 1 Then
            grdOrder.Rows = grdOrder.Rows + 1
            grdOrder.Row = 1
            grdOrder.Col = 1
            grdOrder.SetFocus
        End If
    End If
End Sub

Private Sub grdOrder_EnterCell()
    If grdOrder.TextMatrix(grdOrder.Row, 2) = "" Then
        grdOrder.Col = 1
        Exit Sub
    End If
    grdOrder.Editable = flexEDKbdMouse
    If grdOrder.Col = 2 Then
        grdOrder.Editable = flexEDNone
    End If
    If grdOrder.Col = 4 Then
        grdOrder.Editable = flexEDNone
    End If
    If grdOrder.Col = 7 Then
        grdOrder.Editable = flexEDNone
    End If
End Sub

Private Sub grdOrder_GotFocus()
    cmdPurchase.Visible = True
End Sub

Private Sub grdOrder_KeyDown(KeyCode As Integer, Shift As Integer)
With grdOrder
    Select Case KeyCode
        Case 13 'Enter
            If grdOrder.Col = 1 Then
                Load frmSearch
                frmSearch.Tag = "Order"
                frmSearch.Show vbModal
                Select Case frmSearch.Tag
                    Case ""
                    Case Else
                        grdOrder.TextMatrix(grdOrder.Row, 2) = frmSearch.Tag
                        grdOrder.TextMatrix(grdOrder.Row, 3) = "1"
                        grdOrder.TextMatrix(grdOrder.Row, 4) = "0"
                        grdOrder.TextMatrix(grdOrder.Row, 5) = "0"
                        grdOrder.TextMatrix(grdOrder.Row, 8) = "0.00"
                        grdOrder.TextMatrix(grdOrder.Row, 0) = Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1))
                        ActiveReadServer "Select Pack_Size,Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
                        If rs.RecordCount > 0 Then
                            grdOrder.TextMatrix(grdOrder.Row, 6) = rs.Fields("Landed_Cost")
                            grdOrder.TextMatrix(grdOrder.Row, 7) = rs.Fields("Sales_Tax") & "%"
                            grdOrder.TextMatrix(grdOrder.Row, 9) = rs.Fields("Department_No")
                            grdOrder.TextMatrix(grdOrder.Row, 3) = rs.Fields("Pack_Size")
                        End If
                        rs.Close
                        grdOrder.Col = 5
                        ActiveReadServer2 "Select * from Suppliers_Links_View where Supplier_No = '" & lblGRV.Caption & "' and Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
                        If rs2.RecordCount > 0 Then
                            grdOrder.TextMatrix(grdOrder.Row, 1) = rs2.Fields("Supplier_Code")
                        End If
                        rs2.Close
                        frmMain.Toolbar1.Buttons(2).Enabled = True
                        frmMain.Toolbar1.Buttons(3).Enabled = False
                        frmMain.Toolbar1.Buttons(4).Enabled = True
                        frmSearch.Tag = ""
                End Select
            End If
        Case 106
            frmProdSearch.Show vbModal
            If frmOrder.Tag <> "" Then
                grdOrder.TextMatrix(grdOrder.Row, 1) = frmOrder.Tag
                frmOrder.Tag = ""
                grdOrder.TextMatrix(grdOrder.Row, 3) = "1"
                grdOrder.TextMatrix(grdOrder.Row, 4) = "0"
                grdOrder.TextMatrix(grdOrder.Row, 5) = "0"
                grdOrder.TextMatrix(grdOrder.Row, 0) = Trim(Mid(grdOrder.Text, InStrRev(grdOrder.Text, "-") + 1))
                ActiveReadServer "Select Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    grdOrder.TextMatrix(grdOrder.Row, 6) = rs.Fields("Landed_Cost")
                    grdOrder.TextMatrix(grdOrder.Row, 7) = rs.Fields("Sales_Tax") & "%"
                    grdOrder.TextMatrix(grdOrder.Row, 9) = rs.Fields("Department_No")
                End If
                rs.Close
                grdOrder.Col = 5
                frmMain.Toolbar1.Buttons(2).Enabled = True
                frmMain.Toolbar1.Buttons(3).Enabled = False
                frmMain.Toolbar1.Buttons(4).Enabled = True
            End If
        Case 46
            grdOrder.RemoveItem (grdOrder.Row)
            Calculate_Totals
            If .Rows = 1 Then
                frmMain.Toolbar1.Buttons(2).Enabled = False
                frmMain.Toolbar1.Buttons(3).Enabled = True
                frmMain.Toolbar1.Buttons(4).Enabled = False
            End If
        Case 45, 48 To 57, 96 To 105, 109, 110, 189
            Select Case grdOrder.Col
                Case 2, 3, 4, 7
                Case Else
                    .EditCell
            End Select
        Case 37 'left
            If .Col = 1 Then
                KeyCode = 0
                .Col = 9
            End If
        Case 38 'up
            If grdOrder.Row = 1 Then
                cmbSuppliers.SetFocus
            End If
            If Trim(grdOrder.TextMatrix(grdOrder.Row, 2)) = "" Then
                grdOrder.RemoveItem (grdOrder.Row)
            End If
        Case 39 'Right
            If .Col = 9 Then
                KeyCode = 0
                .Col = 1
            End If
        Case 40 'down
            If Trim(grdOrder.TextMatrix(grdOrder.Row, 2)) = "" Then
                KeyCode = 0
                cmbSuppliers.SetFocus
                If grdOrder.Rows > 2 Then grdOrder.RemoveItem (grdOrder.Row)
            Else
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    grdOrder.ShowCell .Row, 0
                    .Col = 1
                End If
            End If
    End Select
End With
End Sub
Private Sub grdOrder_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(grdOrder.EditText, ".") <> 0 And KeyAscii = 46 Then
        KeyAscii = 0
    End If
    Select Case KeyAscii
        Case 8, 13, 27, 45, 46, 48 To 57
        Case Else
            If Col > 2 Then KeyAscii = 0
    End Select
End Sub
Private Sub grdOrder_LostFocus()
    cmdPurchase.Visible = False
End Sub

Private Sub grdSupp_DblClick()
    cmdSupplier.Value = Up
    cmdSupplier_Click
    txtAddress.Tag = grdSupp.TextMatrix(grdSupp.Row, 0)
    cmbSuppliers.Text = grdSupp.TextMatrix(grdSupp.Row, 1) & " (" & grdSupp.TextMatrix(grdSupp.Row, 0) & ")"
    LoadInfo
    grdOrder_Click
End Sub

Private Sub grdSupp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            KeyCode = 0
            cmdSupplier.Value = Up
            cmdSupplier_Click
            txtAddress.Tag = grdSupp.TextMatrix(grdSupp.Row, 0)
            cmbSuppliers.Text = grdSupp.TextMatrix(grdSupp.Row, 1) & " (" & grdSupp.TextMatrix(grdSupp.Row, 0) & ")"
            LoadInfo
            grdOrder_Click
            KeyCode = 0
        Case 27
            KeyCode = 0
            cmdSupplier.Value = Up
            cmdSupplier_Click
            KeyCode = 0
            cmbSuppliers.SetFocus
    End Select
End Sub
Private Sub LoadInfo()
    On Error Resume Next
    ActiveReadServer "Select * from Suppliers where Supplier_No='" & txtAddress.Tag & "'"
    If rs.RecordCount > 0 Then
        txtAddress.Text = rs.Fields("Address")
        lblContact.Caption = rs.Fields("Contact_Person")
        lblVat.Caption = rs.Fields("VAT_No") & ""
        lblTel.Caption = rs.Fields("Business_Tel") & ""
        lblFax.Caption = rs.Fields("Fax_Tel") & ""
    End If
    rs.Close
    lblGRV.Caption = txtAddress.Tag
    grdOrder.SetFocus
    On Error GoTo 0
End Sub

Private Sub optDate_Click(Index As Integer)
    Select Case Index
        Case 0
            DTStart.Value = DateAdd("d", -7, Date)
            DTStop.Value = Date
        Case 1
            DTStart.Value = DateAdd("M", -1, Date)
            DTStop.Value = Date
    End Select
    GetConsumption
End Sub
