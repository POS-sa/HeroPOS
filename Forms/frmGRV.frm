VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmGRV 
   Caption         =   "Goods Receive Voucher"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   14070
   WindowState     =   2  'Maximized
   Begin btButtonEx.ButtonEx cmdEditgrv 
      Height          =   285
      Left            =   120
      TabIndex        =   68
      ToolTipText     =   " Click to Search and Edit an Old GRV... first select a supplier then click this button."
      Top             =   7680
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Edit GRV"
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
   Begin VB.PictureBox picSundries 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0E8DF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   5760
      ScaleHeight     =   1155
      ScaleWidth      =   3825
      TabIndex        =   54
      Top             =   7590
      Visible         =   0   'False
      Width           =   3825
      Begin VB.OptionButton Option1 
         BackColor       =   &H00F0E8DF&
         Caption         =   "Excl."
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   62
         Top             =   570
         Width           =   705
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00F0E8DF&
         Caption         =   "Incl."
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   61
         Top             =   210
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.TextBox txtTransport 
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
         Left            =   1950
         TabIndex        =   57
         Text            =   "0.00"
         Top             =   810
         Width           =   1725
      End
      Begin VB.TextBox txtUllages 
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
         Left            =   1950
         TabIndex        =   56
         Text            =   "0.00"
         Top             =   450
         Width           =   1725
      End
      Begin VB.TextBox txtDiscount 
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
         Left            =   1950
         TabIndex        =   55
         Text            =   "0.00"
         Top             =   90
         Width           =   1725
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "- Discounts:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   750
         TabIndex        =   60
         Top             =   90
         Width           =   1005
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "- Wastages:"
         Height          =   195
         Left            =   750
         TabIndex        =   59
         Top             =   465
         Width           =   1005
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "+ Transport:"
         Height          =   195
         Left            =   750
         TabIndex        =   58
         Top             =   840
         Width           =   1005
      End
      Begin MSForms.Image Image20 
         Height          =   315
         Index           =   5
         Left            =   1770
         Top             =   60
         Width           =   1995
         BackColor       =   16777215
         Size            =   "3519;556"
      End
      Begin MSForms.Image Image20 
         Height          =   315
         Index           =   4
         Left            =   1770
         Top             =   420
         Width           =   1995
         BackColor       =   16777215
         Size            =   "3519;556"
      End
      Begin MSForms.Image Image20 
         Height          =   315
         Index           =   3
         Left            =   1770
         Top             =   780
         Width           =   1995
         BackColor       =   16777215
         Size            =   "3519;556"
      End
   End
   Begin VB.TextBox txtSubtotal 
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
      TabIndex        =   51
      Text            =   "0.00"
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox txtVat 
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
      TabIndex        =   50
      Text            =   "0.00"
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtTotal 
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
      TabIndex        =   49
      Text            =   "0.00"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox txtInvoice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7950
      TabIndex        =   48
      Top             =   1420
      Width           =   2025
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
      TabIndex        =   46
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
      TabIndex        =   45
      Text            =   "0.00"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1785
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
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   44
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
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   43
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
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "0.00"
      Top             =   7680
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid grdSupp 
      Height          =   150
      Left            =   60
      TabIndex        =   24
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
      FormatString    =   $"frmGRV.frx":0000
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   570
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   73203715
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   11310
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   990
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   73203715
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   345
      Left            =   11310
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1410
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   73203715
      CurrentDate     =   38862
   End
   Begin VSFlex8Ctl.VSFlexGrid grdGRV 
      Height          =   4830
      Left            =   60
      TabIndex        =   2
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
      FormatString    =   $"frmGRV.frx":0078
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   14
         Top             =   5670
         Width           =   1005
      End
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
   End
   Begin btButtonEx.ButtonEx cmdSupplier 
      Height          =   285
      Left            =   180
      TabIndex        =   23
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
      TabIndex        =   1
      Top             =   900
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1349
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmGRV.frx":00F0
   End
   Begin btButtonEx.ButtonEx cmdStrip 
      Height          =   300
      Left            =   12000
      TabIndex        =   41
      ToolTipText     =   " Click to Search.... "
      Top             =   1890
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   4210752
      Caption         =   "Strip Vat from Line"
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
   Begin btButtonEx.ButtonEx cmdSundries 
      Height          =   300
      Left            =   12340
      TabIndex        =   53
      ToolTipText     =   " Click to Search.... "
      Top             =   7200
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      Appearance      =   3
      BorderColor     =   4210752
      Caption         =   "Sundries"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
   Begin btButtonEx.ButtonEx cmdLabel 
      Height          =   300
      Left            =   10500
      TabIndex        =   63
      ToolTipText     =   " Click to Search.... "
      Top             =   1890
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   529
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   4210752
      Caption         =   "Print Labels"
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
   Begin btButtonEx.ButtonEx cmdPrice 
      Height          =   300
      Left            =   9000
      TabIndex        =   64
      ToolTipText     =   " Click to Search.... "
      Top             =   1890
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   529
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   4210752
      Caption         =   "Change Price"
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
   Begin btButtonEx.ButtonEx cmdPdt 
      Height          =   285
      Left            =   3660
      TabIndex        =   65
      ToolTipText     =   " Click to Search.... "
      Top             =   180
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Load Quantities from PDT"
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
   Begin btButtonEx.ButtonEx cmdProducts 
      Height          =   300
      Left            =   120
      TabIndex        =   66
      ToolTipText     =   " Click to Search.... "
      Top             =   7200
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   529
      Appearance      =   3
      BorderColor     =   4210752
      Caption         =   "Edit Product..."
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
   Begin btButtonEx.ButtonEx cmdHistory 
      Height          =   300
      Left            =   2130
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   " Click to Search.... "
      Top             =   7200
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   529
      Appearance      =   3
      BorderColor     =   4210752
      Caption         =   "Purchase History..."
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
      TabIndex        =   47
      Text            =   "0.00"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1785
   End
   Begin btButtonEx.ButtonEx cmdGoodsReturned 
      Height          =   315
      Left            =   1380
      TabIndex        =   36
      ToolTipText     =   " Click to Search.... "
      Top             =   8370
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
   Begin btButtonEx.ButtonEx cmdShortDelivered 
      Height          =   315
      Left            =   1380
      TabIndex        =   34
      ToolTipText     =   " Click to Search.... "
      Top             =   7650
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
   Begin btButtonEx.ButtonEx cmdOverCharged 
      Height          =   315
      Left            =   1380
      TabIndex        =   35
      ToolTipText     =   " Click to Search.... "
      Top             =   8010
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Left            =   5400
      TabIndex        =   52
      Top             =   1920
      Width           =   3555
   End
   Begin MSForms.Image Image20 
      Height          =   315
      Index           =   2
      Left            =   11340
      Top             =   8370
      Width           =   2205
      BackColor       =   16777215
      Size            =   "3889;556"
   End
   Begin MSForms.Image Image20 
      Height          =   315
      Index           =   1
      Left            =   11340
      Top             =   8010
      Width           =   2205
      BackColor       =   16777215
      Size            =   "3889;556"
   End
   Begin MSForms.Image Image20 
      Height          =   315
      Index           =   0
      Left            =   11340
      Top             =   7650
      Width           =   2205
      BackColor       =   16777215
      Size            =   "3889;556"
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
   Begin MSForms.CheckBox chkReturn 
      Height          =   375
      Left            =   4080
      TabIndex        =   40
      Top             =   8340
      Visible         =   0   'False
      Width           =   1335
      BackColor       =   16777215
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2355;661"
      Value           =   "0"
      Caption         =   "Do not Raise"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkOver 
      Height          =   315
      Left            =   4080
      TabIndex        =   39
      Top             =   8010
      Visible         =   0   'False
      Width           =   1335
      BackColor       =   16777215
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2355;556"
      Value           =   "0"
      Caption         =   "Do not Raise"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkShort 
      Height          =   315
      Left            =   4080
      TabIndex        =   38
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
      BackColor       =   16777215
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2355;556"
      Value           =   "0"
      Caption         =   "Do not Raise"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label DTPicker1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11400
      TabIndex        =   37
      Top             =   270
      Width           =   1995
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
      Left            =   210
      TabIndex        =   33
      Top             =   1920
      Width           =   5145
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
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
      Left            =   1980
      TabIndex        =   32
      Top             =   7230
      Width           =   3585
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Totals"
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
      Left            =   9750
      TabIndex        =   31
      Top             =   7250
      Width           =   3765
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
      TabIndex        =   30
      Top             =   7250
      Width           =   3765
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   4
      Left            =   9630
      Top             =   7170
      Width           =   4020
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7091;661"
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   5730
      Top             =   7170
      Width           =   3885
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6853;661"
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
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6250
      TabIndex        =   29
      Top             =   7680
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vat:"
      Height          =   195
      Left            =   6250
      TabIndex        =   28
      Top             =   8055
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   6250
      TabIndex        =   27
      Top             =   8430
      Width           =   1005
   End
   Begin VB.Label lblContact 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7950
      TabIndex        =   25
      Top             =   1020
      Width           =   2025
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
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6210
      TabIndex        =   26
      Top             =   1020
      Width           =   1545
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
   Begin VB.Label lblVat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7950
      TabIndex        =   22
      Top             =   660
      Width           =   1605
   End
   Begin VB.Label lblOrder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7950
      TabIndex        =   20
      Top             =   270
      Width           =   1545
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
      TabIndex        =   21
      Top             =   630
      Width           =   1545
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10260
      TabIndex        =   9
      Top             =   1050
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
      TabIndex        =   19
      Top             =   240
      Width           =   1005
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
      TabIndex        =   18
      Top             =   645
      Width           =   1005
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10260
      TabIndex        =   17
      Top             =   1455
      Width           =   1005
   End
   Begin VB.Label lblGRV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      TabIndex        =   4
      Top             =   225
      Width           =   1935
   End
   Begin MSForms.Image Image7 
      Height          =   285
      Left            =   1440
      Top             =   180
      Width           =   2175
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3836;503"
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   10290
      TabIndex        =   16
      Top             =   8430
      Width           =   1005
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vat:"
      Height          =   195
      Left            =   10290
      TabIndex        =   15
      Top             =   8055
      Width           =   1005
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10290
      TabIndex        =   12
      Top             =   7680
      Width           =   1005
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6210
      TabIndex        =   11
      Top             =   1410
      Width           =   1545
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
      TabIndex        =   10
      Top             =   240
      Width           =   1545
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   165
      Left            =   390
      TabIndex        =   5
      Top             =   900
      Width           =   1005
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GRV Number:"
      Height          =   195
      Left            =   390
      TabIndex        =   3
      Top             =   240
      Width           =   1005
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
   Begin MSForms.Image Image4 
      Height          =   1215
      Left            =   9630
      Top             =   7560
      Width           =   4020
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7091;2143"
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
   Begin MSForms.Image Image9 
      Height          =   345
      Left            =   7800
      Top             =   570
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "3942;609"
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
   Begin MSForms.Image Image13 
      Height          =   315
      Left            =   7320
      Top             =   7650
      Width           =   2205
      BorderColor     =   12632256
      Size            =   "3889;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image15 
      Height          =   315
      Left            =   7320
      Top             =   8370
      Width           =   2205
      BorderColor     =   12632256
      Size            =   "3889;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image14 
      Height          =   315
      Left            =   7320
      Top             =   8010
      Width           =   2205
      BorderColor     =   12632256
      Size            =   "3889;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image5 
      Height          =   1215
      Left            =   5730
      Top             =   7560
      Width           =   3885
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6853;2143"
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
End
Attribute VB_Name = "frmGRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Update_existing_stocklevels()
' Subtract or add to current stocklevels if grv was acually edited
' ???Maybe better rather delete all the old quantities from stock then ad new quantities back in stock


'ActiveUpdateServer " UPDATE QUANTITIES WHERE PRODUCT_NO = THEPRODUCT AND LOCATION = LOCATION SET STOCK_ON_HAND = STOCKONHAND"
'RS.CLOSE
'ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'CHANGED STOCKLEVEL FOR  & PRODUCT_NO AND NAME')"
'RS.CLOSE
End Sub

Public Sub Update_OldGRV()
' after editing
'Change the old grv
'ActiveUpdateServer " DELETE * FROM Purchase_Journal WHERE Supplier_No = Themissingsupplier"
'RS.CLOSE
'ActiveUpdateServer " INSERT INTO Purchase_Journal () VALUES ' ', ' ', ' ', ' '"
'RS.CLOSE



'If txtInvoice.Text = "" And cmbSuppliers.Text <> "" Then
'        MsgBox "Please Enter an Invoice Number", vbCritical, "HeroPOS"
'        txtInvoice.SetFocus
'        Exit Sub
'    End If
'    Select Case MsgBox("Do you want to Save the Edited Goods Receiving Voucher Remember this will change stock quantities?", vbYesNoCancel, "HeroPOS")
'        Case vbYes
'            Screen.MousePointer = 11
'            If Val(lblGRV.Caption) <> 0 Then
'                ActiveUpdateServer "Delete from Purchase_Journal_Listing where GRV_No = " & Val(lblGRV.Caption)
'            Else
'                ActiveReadServer1 "Select isnull(max(GRV_No),0) + 1 as GRV_No from Purchase_Journal"
'                lblGRV.Caption = Format(rs1.Fields("GRV_No"), "000000")
'                rs1.Close
'            End If
'            ActiveUpdateServer "Insert into Purchase_Journal (GRV_No)  values (" & Val(lblGRV.Caption) & ")"
'            DoEvents
'            For i = 1 To grdGRV.Rows - 1
'                ActiveUpdateServer "INSERT INTO Purchase_Journal_Listing(Supplier_No,Workstation_No, User_No, Location_No, GRV_No, Order_No, Invoice_No, Product_Code, Department_No, Pack_Size, Qty_Ordered, Qty_Delivered, Qty_Invoiced, Qty_Returned, Price_Ordered, Price_Invoiced,Vat_Rate, Line_Total,Date_Time,Order_Date,Invoice_Date,Delivery_Date)" & _
'                " VALUES ('" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "'," & Workstation_No & "," & UserRecord.User_Number & "," & Mid(grdGRV.TextMatrix(i, 2), 1, InStr(grdGRV.TextMatrix(i, 2), "-") - 1) & "," & Val(lblGRV.Caption) & "," & Val(lblOrder.Caption) & "," & Val(txtInvoice.Text) & ",'" & grdGRV.TextMatrix(i, 0) & "','" & grdGRV.TextMatrix(i, 10) & "'," & grdGRV.TextMatrix(i, 3) & "," & grdGRV.TextMatrix(i, 4) & "," & grdGRV.TextMatrix(i, 5) & "," & grdGRV.TextMatrix(i, 6) & ",0,0," & grdGRV.ValueMatrix(i, 7) & "," & Replace(grdGRV.TextMatrix(i, 8), "%", "") & "," & grdGRV.ValueMatrix(i, 9) & ",GetDate(),'" & DTPicker2.Value & "','" & DTPicker3.Value & "','" & DTPicker4.Value & "')"
'            Next i
'            MsgBox "Goods Receive Voucher no: " & lblGRV.Caption & " saved Successfully", vbInformation, "HeroPOS"
'            Screen.MousePointer = 0
'        Case vbNo
'            frmMain.picProdBar.Visible = False
'            Unload frmGRV
'            DoEvents
'            frmMain.cmdBar(7).Enabled = True
'            frmDetails.Show
'        Case Else
'    End Select

End Sub
Public Sub List_OldGRVS()
' List all old Grvs as per supplier
'ActiveReadServer " SELECT * FROM Purchase_Journal WHERE Supplier_No = Themissingsupplier"
'RS.CLOSE

'CHANGE THE GRID FOR HEADERS AND INPUT
'grdGRV
'    grdGRV.RowHeight(0) = 550
'    grdGRV.Rows = 1
'    grdGRV.Cols = 11
'    grdGRV.ColHidden(0) = True
'    grdGRV.ColHidden(10) = True
'    grdGRV.TextMatrix(0, 0) = "Product Code"
'    grdGRV.TextMatrix(0, 1) = "Description"
'    grdGRV.TextMatrix(0, 2) = "Receiving Location "
'    grdGRV.TextMatrix(0, 3) = "Pack Size "
'    grdGRV.TextMatrix(0, 4) = "Qty Ordered"
'    grdGRV.TextMatrix(0, 5) = "Qty Delivered"
'    grdGRV.TextMatrix(0, 6) = "Qty Invoiced"
'    grdGRV.TextMatrix(0, 7) = "Unit Price "
'    grdGRV.TextMatrix(0, 8) = "Vat "
'    grdGRV.TextMatrix(0, 9) = "Line Total "
'    grdGRV.ColAlignment(2) = flexAlignLeftCenter
'    grdGRV.ColAlignment(3) = flexAlignRightCenter
'    grdGRV.ColAlignment(4) = flexAlignRightCenter
'    grdGRV.ColAlignment(5) = flexAlignRightCenter
'    grdGRV.ColAlignment(6) = flexAlignRightCenter
'    grdGRV.ColAlignment(7) = flexAlignRightCenter
'    grdGRV.ColAlignment(8) = flexAlignRightCenter
'    grdGRV.ColAlignment(9) = flexAlignRightCenter
'    grdGRV.ColWidth(1) = grdGRV.Width * 0.25
'    grdGRV.ColWidth(2) = grdGRV.Width * 0.18
'    grdGRV.ColWidth(3) = grdGRV.Width * 0.07
'    grdGRV.ColWidth(4) = grdGRV.Width * 0.07
'    grdGRV.ColWidth(5) = grdGRV.Width * 0.07
'    grdGRV.ColWidth(6) = grdGRV.Width * 0.07
'    grdGRV.ColWidth(7) = grdGRV.Width * 0.11
'    grdGRV.ColWidth(8) = grdGRV.Width * 0.08
'    grdGRV.ColWidth(9) = grdGRV.Width * 0.11
'    grdGRV.ColFormat(7) = "0.00"
'    grdGRV.ColFormat(9) = "0.00"
End Sub
Public Sub Load_OldGRV()
'Load the old grv into the page to be viewed and then edited
'    grdGRV.RowHeight(0) = 550
'    grdGRV.Rows = 1
'    grdGRV.Cols = 11
'    grdGRV.ColHidden(0) = True
'    grdGRV.ColHidden(10) = True
'    grdGRV.TextMatrix(0, 0) = "Product Code"
'    grdGRV.TextMatrix(0, 1) = "Description"
'    grdGRV.TextMatrix(0, 2) = "Receiving Location "
'    grdGRV.TextMatrix(0, 3) = "Pack Size "
'    grdGRV.TextMatrix(0, 4) = "Qty Ordered"
'    grdGRV.TextMatrix(0, 5) = "Qty Delivered"
'    grdGRV.TextMatrix(0, 6) = "Qty Invoiced"
'    grdGRV.TextMatrix(0, 7) = "Unit Price "
'    grdGRV.TextMatrix(0, 8) = "Vat "
'    grdGRV.TextMatrix(0, 9) = "Line Total "
'    grdGRV.ColAlignment(2) = flexAlignLeftCenter
'    grdGRV.ColAlignment(3) = flexAlignRightCenter
'    grdGRV.ColAlignment(4) = flexAlignRightCenter
'    grdGRV.ColAlignment(5) = flexAlignRightCenter
'    grdGRV.ColAlignment(6) = flexAlignRightCenter
'    grdGRV.ColAlignment(7) = flexAlignRightCenter
'    grdGRV.ColAlignment(8) = flexAlignRightCenter
'    grdGRV.ColAlignment(9) = flexAlignRightCenter
'    grdGRV.ColWidth(1) = grdGRV.Width * 0.25
'    grdGRV.ColWidth(2) = grdGRV.Width * 0.18
'    grdGRV.ColWidth(3) = grdGRV.Width * 0.07
'    grdGRV.ColWidth(4) = grdGRV.Width * 0.07
'    grdGRV.ColWidth(5) = grdGRV.Width * 0.07
'    grdGRV.ColWidth(6) = grdGRV.Width * 0.07
'    grdGRV.ColWidth(7) = grdGRV.Width * 0.11
'    grdGRV.ColWidth(8) = grdGRV.Width * 0.08
'    grdGRV.ColWidth(9) = grdGRV.Width * 0.11
'    grdGRV.ColFormat(7) = "0.00"
'    grdGRV.ColFormat(9) = "0.00"
'      txtDiscount.Text = "0.00"
'    txtUllages.Text = "0.00"
'    txtTransport.Text = "0.00"
'    ActiveReadServer "Select * from Purchase_Journal_Listing where GRV_No = " & GRV_No
'    If rs.RecordCount > 0 Then
'        lblGRV.Caption = Format(GRV_No, "000000")
'        txtInvoice.Text = rs.Fields("Invoice_No")
'        txtAddress.Tag = rs.Fields("Supplier_No")
'        lblOrder.Caption = Format(rs.Fields("Order_No"), "000000")
'        DTPicker1.Caption = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
'        DTPicker2.Value = rs.Fields("Order_Date")
'        DTPicker3.Value = rs.Fields("Delivery_Date")
'        DTPicker4.Value = rs.Fields("Invoice_Date")
'        ActiveReadServer1 "Select * from Suppliers where Supplier_No='" & txtAddress.Tag & "'"
'        If rs1.RecordCount > 0 Then
'            cmbSuppliers.Text = rs1.Fields("Supplier_Name") & " (" & rs1.Fields("Supplier_No") & ")"
'            txtAddress.Text = rs1.Fields("Address")
'            lblContact.Caption = rs1.Fields("Contact_Person")
'            lblVat.Caption = rs1.Fields("VAT_No") & ""
'        End If
'        rs1.Close
'    End If
'    i = 0
'    grdGRV.Rows = 1
'    While Not rs.EOF
'        i = i + 1
'        grdGRV.Rows = grdGRV.Rows + 1
'        grdGRV.TextMatrix(i, 0) = rs.Fields("Product_Code")
'        ActiveReadServer1 "Select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
'        "+ Unit_of_Measure END AS Description from Products where Product_Code='" & rs.Fields("Product_Code") & "'"
'        If rs1.RecordCount > 0 Then
'            grdGRV.TextMatrix(i, 1) = rs1.Fields("Description") & " - " & rs.Fields("Product_Code")
'        Else
'            grdGRV.TextMatrix(i, 1) = "Unknown Product - " & rs.Fields("Product_Code")
'        End If
'        rs1.Close
'        ActiveReadServer1 "Select Loc_Name from Locations where Location_No=" & rs.Fields("Location_No")
'        If rs1.RecordCount > 0 Then
'            grdGRV.TextMatrix(i, 2) = rs.Fields("Location_No") & " - " & rs1.Fields("Loc_Name")
'        End If
'        rs1.Close
'        grdGRV.TextMatrix(i, 3) = rs.Fields("Pack_Size")
'        grdGRV.TextMatrix(i, 4) = rs.Fields("Qty_Ordered")
'        grdGRV.TextMatrix(i, 5) = rs.Fields("Qty_Delivered")
'        grdGRV.TextMatrix(i, 6) = rs.Fields("Qty_Invoiced")
'        grdGRV.TextMatrix(i, 7) = rs.Fields("Price_Invoiced")
'        grdGRV.TextMatrix(i, 8) = rs.Fields("Vat_Rate") & "%"
'        grdGRV.TextMatrix(i, 10) = rs.Fields("Department_No")
'        grdGRV.TextMatrix(i, 9) = rs.Fields("Line_Total")
'        rs.MoveNext
'    Wend
'    rs.Close
'    Calculate_Totals
'    frmMain.Toolbar1.Buttons(2).Enabled = True
'    frmMain.Toolbar1.Buttons(3).Enabled = False
'    frmMain.Toolbar1.Buttons(4).Enabled = True
End Sub
Public Sub Accept_GRV()
    If txtInvoice.Text = "" And cmbSuppliers.Text <> "" Then
        MsgBox "Please Enter an Invoice Number", vbCritical, "HeroPOS"
        txtInvoice.SetFocus
        Exit Sub
    End If
    frmMain.Toolbar1.Tag = ""
    If Val(txtSubtotal.Text) <> Val(txtCalc1.Text) Then
        MsgBox "Calculated Subtotal do not match the Invoice Subtotal", vbCritical, "HeroPOS"
        txtSubtotal.SetFocus
        Exit Sub
    End If
    If Val(txtTotal.Text) <> Val(txtCalc3.Text) Then
        MsgBox "Calculated Total do not match the Invoice Total", vbCritical, "HeroPOS"
        txtTotal.SetFocus
        Exit Sub
    End If
    If Val(txtVat.Text) <> Val(txtCalc2.Text) Then
        MsgBox "Calculated Vat Total do not match the Invoice Vat Total", vbCritical, "HeroPOS"
        txtVat.SetFocus
        Exit Sub
    End If
    If Val(txtSubtotal.Text) <> Val(txtCalc1.Text) Then
        MsgBox "Calculated Subtotal do not match the Invoice Subtotal", vbCritical, "HeroPOS"
        Exit Sub
    End If
    Select Case MsgBox("Do you want to Accept the Current Goods Receive Voucher?", vbYesNoCancel, "HeroPOS")
        Case vbYes
            Screen.MousePointer = 11
            ActiveUpdateServer "Delete from Purchase_Order_Journal where Order_No = " & Val(lblOrder.Caption)
            If Val(lblGRV.Caption) <> 0 Then
                ActiveUpdateServer "Delete from Purchase_Journal_Listing where GRV_No = " & Val(lblGRV.Caption)
            Else
                ActiveReadServer1 "Select isnull(max(GRV_No),0) + 1 as GRV_No from Purchase_Journal"
                lblGRV.Caption = Format(rs1.Fields("GRV_No"), "000000")
                rs1.Close
            End If
            DoEvents
            NewQuantity = 0
            Balance = 0
            Land_Flag = 0
            ActiveReadServer "Select Balance, Landed_Cost from Suppliers where Supplier_No = '" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "'"
            If rs.RecordCount > 0 Then
                Balance = rs.Fields("Balance")
                Land_Flag = Val(rs.Fields("Landed_Cost") & "")
            End If
            rs.Close
            factor = 1
            For i = 1 To grdGRV.Rows - 1
                If Trim(grdGRV.TextMatrix(i, 0)) <> "" Then
                    If Option1(0).Value = True Then
                        If Val(txtDiscount.Text) + Val(txtUllages.Text) - Val(txtTransport.Text) <> 0 Then
                            factor = 1 - ((Val(txtDiscount.Text) + Val(txtUllages.Text) - Val(txtTransport.Text)) / (Val(txtTotal.Text) + Val(txtDiscount.Text)))
                        End If
                    Else
                        If Val(txtDiscount.Text) + Val(txtUllages.Text) - Val(txtTransport.Text) <> 0 Then
                            factor = 1 - ((Val(txtDiscount.Text) + Val(txtUllages.Text) - Val(txtTransport.Text)) / Val(txtTotal.Text))
                        End If
                    End If
                    NewQuantity = grdGRV.ValueMatrix(i, 3) * grdGRV.ValueMatrix(i, 5)
                    If NewQuantity = 0 Then
                        NewQuantity = grdGRV.ValueMatrix(i, 3) * grdGRV.ValueMatrix(i, 6)
                    End If
                    ActiveUpdateServer "INSERT INTO Purchase_Journal (Supplier_No,Workstation_No, User_No, Location_No, GRV_No, Order_No, Invoice_No, Product_Code, Department_No, Pack_Size, Qty_Ordered, Qty_Delivered, Qty_Invoiced, Qty_Returned, Price_Ordered, Price_Invoiced,Vat_Rate, Line_Total,Date_Time,Order_Date,Invoice_Date,Delivery_Date,Function_Key,Discount,Ullage,Transport)" & _
                    " VALUES ('" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "'," & Workstation_No & "," & UserRecord.User_Number & "," & Mid(grdGRV.TextMatrix(i, 2), 1, InStr(grdGRV.TextMatrix(i, 2), "-") - 1) & "," & Val(lblGRV.Caption) & "," & Val(lblOrder.Caption) & ",'" & txtInvoice.Text & "','" & grdGRV.TextMatrix(i, 0) & "','" & grdGRV.TextMatrix(i, 10) & "'," & grdGRV.TextMatrix(i, 3) & "," & grdGRV.TextMatrix(i, 4) & "," & grdGRV.TextMatrix(i, 5) & "," & grdGRV.TextMatrix(i, 6) & ",0,0," & grdGRV.ValueMatrix(i, 7) * factor & "," & Val(Replace(grdGRV.TextMatrix(i, 8), "%", "")) & "," & grdGRV.ValueMatrix(i, 9) * factor & ",GetDate(),'" & DTPicker2.Value & " " & Time & "','" & DTPicker3.Value & " " & Time & "','" & DTPicker4.Value & " " & Time & "',16,'" & Val(txtDiscount.Text) & "','" & Val(txtUllages.Text) & "','" & Val(txtTransport.Text) & "')"
                    
                    ActiveReadServer2 "Select * from Supplier_Links where Product_Code = '" & grdGRV.TextMatrix(i, 0) & "' and Supplier_No = '" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "'"
                    If rs2.RecordCount = 0 Then
                        ActiveUpdateServer "INSERT INTO [Supplier_Links]([Supplier_No], [Product_Code], [Supplier_Code], [List_Price], [Date_Time])" & _
                        " VALUES('" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "', '" & grdGRV.TextMatrix(i, 0) & "', ''," & grdGRV.ValueMatrix(i, 7) & ",getdate())"
                    Else
                        ActiveUpdateServer "Update [Supplier_Links]" & _
                        " SET [Supplier_No]='" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "', [Product_Code]='" & grdGRV.TextMatrix(i, 0) & "'" & _
                        ", [List_Price]= " & grdGRV.ValueMatrix(i, 7) & ", [Date_Time]= getdate()" & _
                        " WHERE Supplier_No = '" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "' and Product_Code = '" & grdGRV.TextMatrix(i, 0) & "'"
                    End If
                            
                    If NewQuantity <> 0 Then
                        UpdateAveCost grdGRV.ValueMatrix(i, 0), grdGRV.ValueMatrix(i, 7), NewQuantity, Land_Flag
                    End If
                    ActiveReadServer "Select Stock_on_Hand from Quantities where Product_Code = '" & grdGRV.TextMatrix(i, 0) & "' and Location_No = " & Mid(grdGRV.TextMatrix(1, 2), 1, InStr(grdGRV.TextMatrix(i, 2), "-") - 1)
                    If rs.RecordCount > 0 Then
                        ActiveUpdateServer "Update Quantities Set Stock_on_Hand = Stock_on_Hand + " & NewQuantity & " where Product_Code = '" & grdGRV.TextMatrix(i, 0) & "' and Location_No = " & Mid(grdGRV.TextMatrix(i, 2), 1, InStr(grdGRV.TextMatrix(i, 2), "-") - 1)
                    Else
                        ActiveUpdateServer "INSERT INTO Quantities (Product_Code,Location_No,Stock_on_Hand) values ('" & grdGRV.TextMatrix(i, 0) & "'," & Mid(grdGRV.TextMatrix(i, 2), 1, InStr(grdGRV.TextMatrix(i, 2), "-") - 1) & "," & NewQuantity & ")"
                    End If
                rs.Close
                End If
            Next i
            
            ActiveUpdateServer "INSERT INTO [Supplier_Accounts]([User_No],[Date_Time],[Transaction_Type], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Tender_Type],[Ref_No])" & _
            "VALUES(" & UserRecord.User_Number & ",Getdate(),'Supplier Invoice'," & Val(lblGRV.Caption) & ",'" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "',0," & Val(txtTotal.Text) & "," & Balance + (Val(txtTotal.Text)) & ",'','')"

            ActiveUpdateServer "Update Suppliers set Balance=Balance + " & Val(txtTotal.Text) & " where Supplier_No='" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "'"

            Screen.MousePointer = 0
            frmPayment.Show vbModal
            MsgBox "Goods Receive Voucher no: " & lblGRV.Caption & " Accepted Successfully", vbInformation, "HeroPOS"
            Screen.MousePointer = 0
            frmMain.Toolbar1.Tag = lblGRV.Caption
            Screen.MousePointer = 0
            Form_Load
        Case vbNo
            frmMain.picProdBar.Visible = False
            Unload frmGRV
            DoEvents
            frmMain.cmdBar(7).Enabled = True
            frmDetails.Show
        Case Else
    End Select
End Sub
Public Sub OpenGRV()
    On Error Resume Next
    Load frmOpen
    frmOpen.Tag = "GRV"
    DoEvents
    frmOpen.Show vbModal
    txtInvoice.SetFocus
    If frmGRV.grdSupp.Visible = True Then
        grdSupp.Visible = False
        cmbSuppliers.Enabled = False
        cmdSupplier.Enabled = False
        cmdSupplier.Value = Up
        txtInvoice.Enabled = True
        DTPicker1.Enabled = False
        DTPicker2.Enabled = True
        DTPicker3.Enabled = True
        DTPicker4.Enabled = True
        txtSubtotal.Enabled = True
        txtVat.Enabled = True
        txtTotal.Enabled = True
        grdGRV.SetFocus
        DoEvents
    End If
    On Error GoTo 0
End Sub
Public Sub SaveGrv()
    If txtInvoice.Text = "" And cmbSuppliers.Text <> "" Then
        MsgBox "Please Enter an Invoice Number", vbCritical, "HeroPOS"
        txtInvoice.SetFocus
        Exit Sub
    End If
    Select Case MsgBox("Do you want to Save the Current Goods Receive Voucher?", vbYesNoCancel, "HeroPOS")
        Case vbYes
            Screen.MousePointer = 11
            If Val(lblGRV.Caption) <> 0 Then
                ActiveUpdateServer "Delete from Purchase_Journal_Listing where GRV_No = " & Val(lblGRV.Caption)
            Else
                ActiveReadServer1 "Select isnull(max(GRV_No),0) + 1 as GRV_No from Purchase_Journal"
                lblGRV.Caption = Format(rs1.Fields("GRV_No"), "000000")
                rs1.Close
            End If
            ActiveUpdateServer "Insert into Purchase_Journal (GRV_No)  values (" & Val(lblGRV.Caption) & ")"
            DoEvents
            For i = 1 To grdGRV.Rows - 1
                ActiveUpdateServer "INSERT INTO Purchase_Journal_Listing(Supplier_No,Workstation_No, User_No, Location_No, GRV_No, Order_No, Invoice_No, Product_Code, Department_No, Pack_Size, Qty_Ordered, Qty_Delivered, Qty_Invoiced, Qty_Returned, Price_Ordered, Price_Invoiced,Vat_Rate, Line_Total,Date_Time,Order_Date,Invoice_Date,Delivery_Date)" & _
                " VALUES ('" & Replace(Mid(cmbSuppliers, InStrRev(cmbSuppliers, "(") + 1), ")", "") & "'," & Workstation_No & "," & UserRecord.User_Number & "," & Mid(grdGRV.TextMatrix(i, 2), 1, InStr(grdGRV.TextMatrix(i, 2), "-") - 1) & "," & Val(lblGRV.Caption) & "," & Val(lblOrder.Caption) & "," & Val(txtInvoice.Text) & ",'" & grdGRV.TextMatrix(i, 0) & "','" & grdGRV.TextMatrix(i, 10) & "'," & grdGRV.TextMatrix(i, 3) & "," & grdGRV.TextMatrix(i, 4) & "," & grdGRV.TextMatrix(i, 5) & "," & grdGRV.TextMatrix(i, 6) & ",0,0," & grdGRV.ValueMatrix(i, 7) & "," & Replace(grdGRV.TextMatrix(i, 8), "%", "") & "," & grdGRV.ValueMatrix(i, 9) & ",GetDate(),'" & DTPicker2.Value & "','" & DTPicker3.Value & "','" & DTPicker4.Value & "')"
            Next i
            MsgBox "Goods Receive Voucher no: " & lblGRV.Caption & " saved Successfully", vbInformation, "HeroPOS"
            Screen.MousePointer = 0
        Case vbNo
            frmMain.picProdBar.Visible = False
            Unload frmGRV
            DoEvents
            frmMain.cmdBar(7).Enabled = True
            frmDetails.Show
        Case Else
    End Select
End Sub
Public Sub DeleteGRV()
    Select Case MsgBox("Are you sure you want to Delete the Current Goods Receive Voucher", vbYesNo, "HeroPOS")
        Case vbYes
            ActiveUpdateServer "Delete from Purchase_Journal_Listing where GRV_No = " & Val(lblGRV.Caption)
            ActiveUpdateServer "Delete from Purchase_Journal where GRV_No = " & Val(lblGRV.Caption)
            frmMain.picProdBar.Visible = False
            Unload frmGRV
            DoEvents
            frmMain.cmdBar(7).Enabled = True
            frmDetails.Show
        Case Else
    End Select
End Sub

Private Sub chkOver_Click()
    cmdStrip.Enabled = False
End Sub

Private Sub chkReturn_Click()
    cmdStrip.Enabled = False
End Sub

Private Sub chkShort_Click()
    cmdStrip.Enabled = False
End Sub

Private Sub cmbSuppliers_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbSuppliers_GotFocus()
    cmdStrip.Enabled = False
End Sub

Private Sub cmbSuppliers_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            Else
                txtInvoice.SetFocus
            End If
        Case 38
            DoEvents
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                If grdGRV.Rows = 1 Then
                    txtInvoice.SetFocus
                Else
                    txtTotal.SetFocus
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
    End If
End Sub

Private Sub cmdEditgrv_Click()
ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Grv edit')"

Update_OldGRV
End Sub

Private Sub cmdHistory_Click()
    If Trim(grdGRV.TextMatrix(grdGRV.Row, 0)) = "" Then Exit Sub
    frmGRV.Tag = "1"
    Load frmPurchase
    frmPurchase.Tag = "GRV"
    DoEvents
    frmPurchase.Show vbModal
End Sub
Private Sub cmdLabel_Click()
    frmGRV.Tag = "1"
    Load frmLabel
    frmLabel.Tag = "GRV"
    DoEvents
    frmLabel.Show vbModal
End Sub

Private Sub cmdPdt_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    stime = 0
    frmGRV.Tag = "1"
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
        If grdGRV.FindRow(Barcode, 0, 0) <> -1 Then
            grdGRV.TextMatrix(grdGRV.FindRow(Barcode, 0, 0), 5) = grdGRV.ValueMatrix(grdGRV.FindRow(Barcode, 0, 0), 5) + Qty
        End If
    Wend
    Close filenum
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
trap:
    Close filenum
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
End Sub
Private Sub cmdPrice_Click()
    frmGRV.Tag = "1"
    Load frmPriceChange
    frmPriceChange.Tag = "GRV"
    DoEvents
    frmPriceChange.Show vbModal
End Sub

Private Sub cmdProducts_Click()
    grdGRV.SetFocus
    Load frmProd
    frmGRV.Tag = "1"
    frmProd.Show vbModal
    frmGRV.cmdProducts.Tag = "C"
    grdGRV.SetFocus
End Sub
Private Sub cmdStrip_Click()
    grdGRV.TextMatrix(grdGRV.Row, 7) = grdGRV.TextMatrix(grdGRV.Row, 7) / ((100 + Val(Replace(grdGRV.TextMatrix(grdGRV.Row, 8), "%", ""))) / 100)
    grdGRV.TextMatrix(grdGRV.Row, 9) = grdGRV.ValueMatrix(grdGRV.Row, 6) * grdGRV.ValueMatrix(grdGRV.Row, 7)
    Calculate_Totals
    grdGRV.SetFocus
End Sub
Private Sub cmdSundries_Click()
    Select Case cmdSundries.Value
        Case Up
            picSundries.Visible = False
            Label18 = "Calculated Totals"
        Case Down
            picSundries.Visible = True
            Label18 = "Sundries Charges"
    End Select
    Calculate_Totals
End Sub
Public Sub cmdSupplier_Click()
    Select Case cmdSupplier.Value
        Case 0
            Screen.MousePointer = 11
            grdSupp.Visible = True
            LoadSuppliers
            txtInvoice.Enabled = False
            DTPicker1.Enabled = False
            DTPicker2.Enabled = False
            DTPicker3.Enabled = False
            DTPicker4.Enabled = False
            cmbSuppliers.Enabled = False
            txtSubtotal.Enabled = False
            txtVat.Enabled = False
            txtTotal.Enabled = False
            DoEvents
            If grdSupp.Rows > 1 Then
                grdSupp.Col = 1
                grdSupp.SetFocus
            End If
            Screen.MousePointer = 0
        Case 1
            grdSupp.Visible = False
            cmbSuppliers.Enabled = True
            txtInvoice.Enabled = True
            DTPicker1.Enabled = True
            DTPicker2.Enabled = True
            DTPicker3.Enabled = True
            DTPicker4.Enabled = True
            txtSubtotal.Enabled = True
            txtVat.Enabled = True
            txtTotal.Enabled = True
            DoEvents
    End Select
End Sub
Public Sub LoadOrder(Order_No)
    ActiveReadServer "Select * from Purchase_Order_Journal where Workstation_No is not null and Order_No = " & Order_No
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
        End If
        rs1.Close
    End If
    i = 0
    grdGRV.Rows = 1
    While Not rs.EOF
        i = i + 1
        grdGRV.Rows = grdGRV.Rows + 1
        grdGRV.TextMatrix(i, 0) = rs.Fields("Product_Code")
        ActiveReadServer1 "Select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
        "+ Unit_of_Measure END AS Description from Products where Product_Code='" & rs.Fields("Product_Code") & "'"
        If rs1.RecordCount > 0 Then
         
            grdGRV.TextMatrix(i, 1) = rs1.Fields("Description") & " - " & rs.Fields("Product_Code")
        Else
            grdGRV.TextMatrix(i, 1) = "Unknown Product - " & rs.Fields("Product_Code")
        End If
        rs1.Close
        ActiveReadServer1 "Select Loc_Name from Locations where Location_No=" & rs.Fields("Location_No")
        If rs1.RecordCount > 0 Then
            grdGRV.TextMatrix(i, 2) = rs.Fields("Location_No") & " - " & rs1.Fields("Loc_Name")
        End If
        rs1.Close
        grdGRV.TextMatrix(i, 3) = "1"
        grdGRV.TextMatrix(i, 4) = Round(rs.Fields("Qty_Ordered") * rs.Fields("Pack_Size"), 3)
        grdGRV.TextMatrix(i, 5) = "0"
        grdGRV.TextMatrix(i, 6) = "0"
        If rs.Fields("Price_Ordered") And rs.Fields("Pack_Size") <> 0 Then
        grdGRV.TextMatrix(i, 7) = Format(rs.Fields("Price_Ordered") / rs.Fields("Pack_Size"), "0.00")
        End If
        grdGRV.TextMatrix(i, 8) = rs.Fields("Vat_Rate") & "%"
        grdGRV.TextMatrix(i, 10) = rs.Fields("Department_No")
        grdGRV.TextMatrix(i, 9) = Format(rs.Fields("Line_Total"), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    Calculate_Totals
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(3).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = True
End Sub

Public Sub LoadGrv(GRV_No)
    txtDiscount.Text = "0.00"
    txtUllages.Text = "0.00"
    txtTransport.Text = "0.00"
    ActiveReadServer "Select * from Purchase_Journal_Listing where GRV_No = " & GRV_No
    If rs.RecordCount > 0 Then
        lblGRV.Caption = Format(GRV_No, "000000")
        txtInvoice.Text = rs.Fields("Invoice_No")
        txtAddress.Tag = rs.Fields("Supplier_No")
        lblOrder.Caption = Format(rs.Fields("Order_No"), "000000")
        DTPicker1.Caption = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
        DTPicker2.Value = rs.Fields("Order_Date")
        DTPicker3.Value = rs.Fields("Delivery_Date")
        DTPicker4.Value = rs.Fields("Invoice_Date")
        ActiveReadServer1 "Select * from Suppliers where Supplier_No='" & txtAddress.Tag & "'"
        If rs1.RecordCount > 0 Then
            cmbSuppliers.Text = rs1.Fields("Supplier_Name") & " (" & rs1.Fields("Supplier_No") & ")"
            txtAddress.Text = rs1.Fields("Address")
            lblContact.Caption = rs1.Fields("Contact_Person")
            lblVat.Caption = rs1.Fields("VAT_No") & ""
        End If
        rs1.Close
    End If
    i = 0
    grdGRV.Rows = 1
    While Not rs.EOF
        i = i + 1
        grdGRV.Rows = grdGRV.Rows + 1
        grdGRV.TextMatrix(i, 0) = rs.Fields("Product_Code")
        ActiveReadServer1 "Select CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
        "+ Unit_of_Measure END AS Description from Products where Product_Code='" & rs.Fields("Product_Code") & "'"
        If rs1.RecordCount > 0 Then
            grdGRV.TextMatrix(i, 1) = rs1.Fields("Description") & " - " & rs.Fields("Product_Code")
        Else
            grdGRV.TextMatrix(i, 1) = "Unknown Product - " & rs.Fields("Product_Code")
        End If
        rs1.Close
        ActiveReadServer1 "Select Loc_Name from Locations where Location_No=" & rs.Fields("Location_No")
        If rs1.RecordCount > 0 Then
            grdGRV.TextMatrix(i, 2) = rs.Fields("Location_No") & " - " & rs1.Fields("Loc_Name")
        End If
        rs1.Close
        grdGRV.TextMatrix(i, 3) = rs.Fields("Pack_Size")
        grdGRV.TextMatrix(i, 4) = rs.Fields("Qty_Ordered")
        grdGRV.TextMatrix(i, 5) = rs.Fields("Qty_Delivered")
        grdGRV.TextMatrix(i, 6) = rs.Fields("Qty_Invoiced")
        grdGRV.TextMatrix(i, 7) = rs.Fields("Price_Invoiced")
        grdGRV.TextMatrix(i, 8) = rs.Fields("Vat_Rate") & "%"
        grdGRV.TextMatrix(i, 10) = rs.Fields("Department_No")
        grdGRV.TextMatrix(i, 9) = rs.Fields("Line_Total")
        rs.MoveNext
    Wend
    rs.Close
    Calculate_Totals
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(3).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = True
End Sub

Private Sub DTPicker2_GotFocus()
    cmdStrip.Enabled = False
End Sub

Private Sub DTPicker3_GotFocus()
    cmdStrip.Enabled = False
End Sub

Private Sub DTPicker4_GotFocus()
    cmdStrip.Enabled = False
End Sub

Private Sub Form_Activate()
    If frmGRV.Tag <> "" Then Exit Sub
    frmMain.Toolbar1.Buttons(11).Enabled = False
    frmMain.Toolbar1.Buttons(12).Enabled = False
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Unload rptGRV
    lblAve.Caption = ""
    DTPicker1.Caption = Format(Date, "ddd dd MMM yyyy")
    DTPicker2.Value = Date
    DTPicker3.Value = Date
    DTPicker4.Value = Date
    grdSupp.Height = 4830
    grdGRV.RowHeight(0) = 550
    grdGRV.Rows = 1
    grdGRV.Cols = 11
    grdGRV.ColHidden(0) = True
    grdGRV.ColHidden(10) = True
    grdGRV.TextMatrix(0, 0) = "Product Code"
    grdGRV.TextMatrix(0, 1) = "Description"
    grdGRV.TextMatrix(0, 2) = "Receiving Location "
    grdGRV.TextMatrix(0, 3) = "Pack Size "
    grdGRV.TextMatrix(0, 4) = "Qty Ordered"
    grdGRV.TextMatrix(0, 5) = "Qty Delivered"
    grdGRV.TextMatrix(0, 6) = "Qty Invoiced"
    grdGRV.TextMatrix(0, 7) = "Unit Price "
    grdGRV.TextMatrix(0, 8) = "Vat "
    grdGRV.TextMatrix(0, 9) = "Line Total "
    grdGRV.ColAlignment(2) = flexAlignLeftCenter
    grdGRV.ColAlignment(3) = flexAlignRightCenter
    grdGRV.ColAlignment(4) = flexAlignRightCenter
    grdGRV.ColAlignment(5) = flexAlignRightCenter
    grdGRV.ColAlignment(6) = flexAlignRightCenter
    grdGRV.ColAlignment(7) = flexAlignRightCenter
    grdGRV.ColAlignment(8) = flexAlignRightCenter
    grdGRV.ColAlignment(9) = flexAlignRightCenter
    grdGRV.ColWidth(1) = grdGRV.Width * 0.25
    grdGRV.ColWidth(2) = grdGRV.Width * 0.18
    grdGRV.ColWidth(3) = grdGRV.Width * 0.07
    grdGRV.ColWidth(4) = grdGRV.Width * 0.07
    grdGRV.ColWidth(5) = grdGRV.Width * 0.07
    grdGRV.ColWidth(6) = grdGRV.Width * 0.07
    grdGRV.ColWidth(7) = grdGRV.Width * 0.11
    grdGRV.ColWidth(8) = grdGRV.Width * 0.08
    grdGRV.ColWidth(9) = grdGRV.Width * 0.11
    grdGRV.ColFormat(7) = "0.00"
    grdGRV.ColFormat(9) = "0.00"
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
    lblGRV.Caption = "000000"
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
    frmMain.Toolbar1.Buttons(2).Caption = "Accept"
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

Private Sub grdGRV_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Select Case Col
        Case 1
            If grdGRV.Text <> "" Then
                grdGRV.TextMatrix(grdGRV.Row, 3) = "1"
                grdGRV.TextMatrix(grdGRV.Row, 4) = "0"
                grdGRV.TextMatrix(grdGRV.Row, 5) = "0"
                grdGRV.TextMatrix(grdGRV.Row, 6) = "0"
                grdGRV.TextMatrix(grdGRV.Row, 0) = Trim(Mid(grdGRV.Text, InStrRev(grdGRV.Text, "-") + 1))
                ActiveReadServer "Select Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdGRV.TextMatrix(grdGRV.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    grdGRV.TextMatrix(grdGRV.Row, 7) = rs.Fields("Landed_Cost")
                    grdGRV.TextMatrix(grdGRV.Row, 8) = rs.Fields("Sales_Tax") & "%"
                    grdGRV.TextMatrix(grdGRV.Row, 10) = rs.Fields("Department_No")
                    grdGRV.TextMatrix(grdGRV.Row, 9) = 0
                End If
                rs.Close
                grdGRV.Col = 5
                frmMain.Toolbar1.Buttons(2).Enabled = True
                frmMain.Toolbar1.Buttons(3).Enabled = False
                frmMain.Toolbar1.Buttons(4).Enabled = True
            End If
        Case 2
            grdGRV.Col = 1
            Screen.MousePointer = 11
            grdGRV.ColComboList(1) = ""
            If grdGRV.TextMatrix(grdGRV.Row, 2) = "" Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            Screen.MousePointer = 0
        Case 5
            If grdGRV.ValueMatrix(grdGRV.Row, 0) <> 0 Then
                ActiveReadServer "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdGRV.TextMatrix(grdGRV.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    AveCost = Val(rs.Fields("Ave_Cost") & "")
                    SOH = Val(rs.Fields("SOH") & "")
                    lblInfo.Caption = "Last Landed Cost: " & Format(rs.Fields("Landed_Cost"), "0.00") & "   Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs.Fields("Selling_Price"), "0.00")
                End If
                rs.Close
                
                NEWQTY = grdGRV.ValueMatrix(grdGRV.Row, 6)
                Land = grdGRV.ValueMatrix(grdGRV.Row, 7)
                If SOH < 0 Then SOH = 0
                   
                If NEWQTY <> 0 And Land <> 0 Then
                    NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                    If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                        lblAve.ForeColor = &HC0&
                        lblAve.Caption = "New Average Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    Else
                        lblAve.ForeColor = &HC00000
                        lblAve.Caption = "New Average Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    End If
                Else
                    lblAve.Caption = "New Average Cost = " & Format(Land, "0.00")
                End If
            End If
        Case 6
            grdGRV.TextMatrix(grdGRV.Row, 9) = grdGRV.ValueMatrix(grdGRV.Row, 6) * grdGRV.ValueMatrix(grdGRV.Row, 7)
        Case 7
            If grdGRV.ValueMatrix(grdGRV.Row, 0) <> 0 Then
                grdGRV.TextMatrix(grdGRV.Row, 9) = grdGRV.ValueMatrix(grdGRV.Row, 6) * grdGRV.ValueMatrix(grdGRV.Row, 7)
                ActiveReadServer "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdGRV.TextMatrix(grdGRV.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    AveCost = Val(rs.Fields("Ave_Cost") & "")
                    SOH = Val(rs.Fields("SOH") & "")
                    lblInfo.Caption = "Last Landed Cost: " & Format(rs.Fields("Landed_Cost"), "0.00") & "   Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs.Fields("Selling_Price"), "0.00")
                End If
                rs.Close
                
                NEWQTY = grdGRV.ValueMatrix(grdGRV.Row, 6)
                Land = grdGRV.ValueMatrix(grdGRV.Row, 7)
                If SOH < 0 Then SOH = 0
                   
                If NEWQTY <> 0 And Land <> 0 Then
                    NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                    If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                        lblAve.ForeColor = &HC0&
                        lblAve.Caption = "New Average Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    Else
                        lblAve.ForeColor = &HC00000
                        lblAve.Caption = "New Average Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    End If
                Else
                    lblAve.Caption = "New Average Cost = " & Format(Land, "0.00")
                End If
            End If
            If grdGRV.ValueMatrix(grdGRV.Row, 6) <> 0 Then
                grdGRV.TextMatrix(grdGRV.Row, 9) = grdGRV.ValueMatrix(grdGRV.Row, 6) * grdGRV.ValueMatrix(grdGRV.Row, 7)
            End If
        Case 9
            If grdGRV.ValueMatrix(grdGRV.Row, 6) <> 0 Then
                grdGRV.TextMatrix(grdGRV.Row, 7) = grdGRV.ValueMatrix(grdGRV.Row, 9) / grdGRV.ValueMatrix(grdGRV.Row, 6)
            End If
            If grdGRV.ValueMatrix(grdGRV.Row, 0) <> 0 Then
                grdGRV.TextMatrix(grdGRV.Row, 9) = grdGRV.ValueMatrix(grdGRV.Row, 6) * grdGRV.ValueMatrix(grdGRV.Row, 7)
                ActiveReadServer "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
                "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdGRV.TextMatrix(grdGRV.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    AveCost = Val(rs.Fields("Ave_Cost") & "")
                    SOH = Val(rs.Fields("SOH") & "")
                    lblInfo.Caption = "Last Landed Cost: " & Format(rs.Fields("Landed_Cost"), "0.00") & "   Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs.Fields("Selling_Price"), "0.00")
                End If
                rs.Close
                
                NEWQTY = grdGRV.ValueMatrix(grdGRV.Row, 6)
                Land = grdGRV.ValueMatrix(grdGRV.Row, 7)
                If SOH < 0 Then SOH = 0
                If NEWQTY <> 0 And Land <> 0 Then
                    NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                    If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                        lblAve.ForeColor = &HC0&
                        lblAve.Caption = "New Average Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    Else
                        lblAve.ForeColor = &HC00000
                        lblAve.Caption = "New Average Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    End If
                Else
                    lblAve.Caption = "New Average Cost = " & Format(Land, "0.00")
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
    For i = 1 To grdGRV.Rows - 1
        txtCalc1.Text = Format(Val(txtCalc1.Text) + grdGRV.ValueMatrix(i, 9), "0.00")
        txtCalc3.Text = Val(txtCalc3.Text) + (grdGRV.ValueMatrix(i, 9) * (1 + grdGRV.ValueMatrix(i, 8)))
    Next i
    
    If Option1(0).Value = True Then
        txtCalc1.Text = Format(Val(txtCalc1.Text) - ((Val(txtDiscount.Text)) / 1.14), "0.00")
        txtCalc1.Text = Format(Val(txtCalc1.Text) - ((Val(txtUllages.Text)) / 1.14), "0.00")
        txtCalc1.Text = Format(Val(txtCalc1.Text) + ((Val(txtTransport.Text)) / 1.14), "0.00")
        txtCalc3.Text = Format(Val(txtCalc3.Text) - Val(txtDiscount.Text), "0.00")
        txtCalc3.Text = Format(Val(txtCalc3.Text) - Val(txtUllages.Text), "0.00")
        txtCalc3.Text = Format(Val(txtCalc3.Text) + Val(txtTransport.Text), "0.00")
        txtCalc2.Text = Format(Val(txtCalc3.Text) - Val(txtCalc1.Text), "0.00")
    Else
        txtCalc1.Text = Format(Val(txtCalc1.Text) - Val(txtDiscount.Text), "0.00")
        txtCalc1.Text = Format(Val(txtCalc1.Text) - Val(txtUllages.Text), "0.00")
        txtCalc1.Text = Format(Val(txtCalc1.Text) + Val(txtTransport.Text), "0.00")
        txtCalc3.Text = Format(Val(txtCalc3.Text) - Val(txtDiscount.Text) * 1.14, "0.00")
        txtCalc3.Text = Format(Val(txtCalc3.Text) - Val(txtUllages.Text) * 1.14, "0.00")
        txtCalc3.Text = Format(Val(txtCalc3.Text) + Val(txtTransport.Text) * 1.14, "0.00")
        txtCalc2.Text = Format(Val(txtCalc3.Text) - Val(txtCalc1.Text), "0.00")
    End If
End Sub
Private Sub grdGRV_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    cmdPdt.Enabled = True
    If grdGRV.TextMatrix(NewRow, 0) <> "" And NewRow <> 0 Then
        If grdGRV.ValueMatrix(NewRow, 0) <> 0 Then
            ActiveReadServer "SELECT Products.Ave_Cost, Products.Landed_Cost, Products.Selling_Price," & _
            "(Select sum(isnull(Stock_on_Hand,0)) from Quantities where Products.Product_Code=Quantities.Product_Code) as SOH FROM Products where Products.Product_Code = '" & grdGRV.TextMatrix(NewRow, 0) & "'"
            If rs.RecordCount > 0 Then
                AveCost = Val(rs.Fields("Ave_Cost") & "")
                SOH = Val(rs.Fields("SOH") & "")
                lblInfo.Caption = "Last Landed Cost: " & Format(rs.Fields("Landed_Cost"), "0.00") & "   Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & "   Selling Price: " & Format(rs.Fields("Selling_Price"), "0.00")
                lblInfo.Tag = "Last Landed Cost: " & Format(rs.Fields("Landed_Cost"), "0.00") & Chr(13) & "Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & Chr(13) & "Selling Price: " & Format(rs.Fields("Selling_Price"), "0.00")
            End If
            rs.Close
            
            NEWQTY = grdGRV.ValueMatrix(NewRow, 6)
            Land = grdGRV.ValueMatrix(NewRow, 7)
            If SOH < 0 Then SOH = 0
               
            If NEWQTY <> 0 And Land <> 0 Then
                If SOH + NEWQTY <> 0 Then
                    NewAverage = ((SOH * AveCost) + (NEWQTY * Land)) / (SOH + NEWQTY)
                    If ((100 - Abs((AveCost / NewAverage) * 100)) > 15 Or (100 - Abs((AveCost / NewAverage) * 100)) < -15) And (100 - Abs((AveCost / NewAverage) * 100)) <> 0 Then
                        lblAve.ForeColor = &HC0&
                        lblAve.Caption = "New Average Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                        lblAve.Tag = "New Average Cost = " & Format(NewAverage, "0.00") & Chr(13) & "(" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    Else
                        lblAve.ForeColor = &HC00000
                        lblAve.Caption = "New Average Cost = " & Format(NewAverage, "0.00") & " (" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                        lblAve.Tag = "New Average Cost = " & Format(NewAverage, "0.00") & Chr(13) & "(" & Round((100 - Abs((AveCost / NewAverage) * 100)), 2) & "% Deviation)"
                    End If
                End If
            Else
                lblAve.Caption = "New Average Cost = " & Format(Land, "0.00")
                lblAve.Tag = "New Average Cost = " & Format(Land, "0.00")
            End If
        End If
    End If
    If grdGRV.ValueMatrix(NewRow, 6) > 0 Then
        cmdLabel.Enabled = True
        cmdPrice.Enabled = True
    Else
        cmdLabel.Enabled = False
        cmdPrice.Enabled = False
    End If
    If grdGRV.ValueMatrix(NewRow, 5) > 0 Then
        cmdLabel.Enabled = True
        cmdPrice.Enabled = True
    Else
        cmdLabel.Enabled = False
        cmdPrice.Enabled = False
    End If
    If grdGRV.TextMatrix(NewRow, 0) = "" Then
        cmdProducts.Caption = "Create Product..."
        lblInfo.Visible = False
    Else
        cmdProducts.Caption = "Edit Product..."
        lblInfo.Visible = True
    End If
End Sub

Private Sub grdGRV_Click()
    If txtInvoice.Text <> "" And cmbSuppliers.Text <> "" Then
        If grdGRV.Rows = 1 Then
            grdGRV.Rows = grdGRV.Rows + 1
            grdGRV.Row = 1
            grdGRV.Col = 2
        End If
    End If
End Sub

Private Sub grdGRV_EnterCell()
    If grdGRV.TextMatrix(grdGRV.Row, 2) = "" Then
        grdGRV.Col = 2
    End If
    If grdGRV.Col = 1 Then
        grdGRV.Editable = flexEDNone
    End If
    If grdGRV.Col = 2 And grdGRV.ColComboList(2) = "" Then
        grdGRV.ColComboList(2) = ""
        ActiveReadServer "Select Location_No,Loc_Name from Locations where Loc_Type <> 3 order by Location_no"
        While Not rs.EOF
            grdGRV.ColComboList(2) = grdGRV.ColComboList(2) & rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name") & "|"
            rs.MoveNext
        Wend
        rs.Close
    End If
    If grdGRV.Rows > 1 And grdGRV.Col = 7 Then
        cmdStrip.Enabled = True
    Else
        cmdStrip.Enabled = False
    End If
    If grdGRV.Col = 8 Then
        grdGRV.Editable = flexEDNone
    End If
    If grdGRV.Col = 2 Then
        grdGRV.Editable = flexEDKbdMouse
    End If
    
End Sub

Private Sub grdGRV_GotFocus()
    cmdProducts.Visible = True
    cmdHistory.Visible = True
End Sub

Private Sub grdGRV_KeyDown(KeyCode As Integer, Shift As Integer)
With grdGRV
    Select Case KeyCode
        Case 13 'Enter
            If grdGRV.Col = 1 Then
                Load frmSearch
                frmSearch.Tag = ""
                frmSearch.Show vbModal
                Select Case frmSearch.Tag
                    Case ""
                    Case Else
                        grdGRV.TextMatrix(grdGRV.Row, 1) = frmSearch.Tag
                        grdGRV.TextMatrix(grdGRV.Row, 3) = "1"
                        grdGRV.TextMatrix(grdGRV.Row, 4) = "0"
                        grdGRV.TextMatrix(grdGRV.Row, 5) = "0"
                        grdGRV.TextMatrix(grdGRV.Row, 6) = "0"
                        grdGRV.TextMatrix(grdGRV.Row, 0) = Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1))
                        ActiveReadServer "Select Recipe_Item,Returnable_Item,Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdGRV.TextMatrix(grdGRV.Row, 0) & "'"
                        If rs.RecordCount > 0 Then
                            If Val(rs.Fields("Recipe_Item") & "") = 1 And Val(rs.Fields("Returnable_Item") & "") = 1 Then
                                ActiveReadServer2 "Select * from Recipes where Product_Code = '" & Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1)) & "'"
                                While Not rs2.EOF
                                    grdGRV.TextMatrix(grdGRV.Row, 0) = rs2.Fields("Line_Code")
                                    grdGRV.TextMatrix(grdGRV.Row, 1) = Trim(Mid(rs2.Fields("Description"), 1, InStrRev(rs2.Fields("Description"), ",") - 1)) & " - " & Trim(Mid(rs2.Fields("Description"), InStrRev(rs2.Fields("Description"), ",") + 1))
                                    grdGRV.TextMatrix(grdGRV.Row, 3) = "1"
                                    grdGRV.TextMatrix(grdGRV.Row, 4) = "0"
                                    grdGRV.TextMatrix(grdGRV.Row, 5) = rs2.Fields("Qty_Used")
                                    grdGRV.TextMatrix(grdGRV.Row, 6) = rs2.Fields("Qty_Used")
                                    grdGRV.TextMatrix(grdGRV.Row, 7) = rs2.Fields("Cost")
                                    grdGRV.TextMatrix(grdGRV.Row, 8) = rs.Fields("Sales_Tax") & "%"
                                    grdGRV.TextMatrix(grdGRV.Row, 10) = rs.Fields("Department_No")
                                    grdGRV.TextMatrix(grdGRV.Row, 9) = grdGRV.ValueMatrix(grdGRV.Row, 6) * grdGRV.ValueMatrix(grdGRV.Row, 7)
                                    rs2.MoveNext
                                    If rs2.EOF = False Then
                                        grdGRV.Rows = grdGRV.Rows + 1
                                        grdGRV.Row = grdGRV.Rows - 1
                                        grdGRV.TextMatrix(grdGRV.Row, 2) = grdGRV.TextMatrix(grdGRV.Row - 1, 2)
                                    End If
                                Wend
                                rs2.Close
                            Else
                                grdGRV.TextMatrix(grdGRV.Row, 7) = rs.Fields("Landed_Cost")
                                grdGRV.TextMatrix(grdGRV.Row, 8) = rs.Fields("Sales_Tax") & "%"
                                grdGRV.TextMatrix(grdGRV.Row, 10) = rs.Fields("Department_No")
                                grdGRV.TextMatrix(grdGRV.Row, 9) = 0
                            End If
                        End If
                        rs.Close
                        grdGRV.Col = 5
                        frmMain.Toolbar1.Buttons(2).Enabled = True
                        frmMain.Toolbar1.Buttons(3).Enabled = False
                        frmMain.Toolbar1.Buttons(4).Enabled = True
                        frmSearch.Tag = ""
                End Select
            End If
            If grdGRV.Col = 2 And grdGRV.ColComboList(2) <> "" Then
                grdGRV.ColComboList(2) = ""
                ActiveReadServer "Select Location_No,Loc_Name from Locations where Loc_Type <> 3 order by Location_no"
                While Not rs.EOF
                    grdGRV.ColComboList(2) = grdGRV.ColComboList(2) & rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name") & "|"
                    rs.MoveNext
                Wend
                rs.Close
                grdGRV.Editable = flexEDKbdMouse
            End If
        Case 106
            frmProdSearch.Show vbModal
            If frmGRV.Tag <> "" Then
                grdGRV.TextMatrix(grdGRV.Row, 1) = frmGRV.Tag
                frmGRV.Tag = ""
                grdGRV.TextMatrix(grdGRV.Row, 3) = "1"
                grdGRV.TextMatrix(grdGRV.Row, 4) = "0"
                grdGRV.TextMatrix(grdGRV.Row, 5) = "0"
                grdGRV.TextMatrix(grdGRV.Row, 6) = "0"
                grdGRV.TextMatrix(grdGRV.Row, 0) = Trim(Mid(grdGRV.Text, InStrRev(grdGRV.Text, "-") + 1))
                ActiveReadServer "Select Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdGRV.TextMatrix(grdGRV.Row, 0) & "'"
                If rs.RecordCount > 0 Then
                    grdGRV.TextMatrix(grdGRV.Row, 7) = rs.Fields("Landed_Cost")
                    grdGRV.TextMatrix(grdGRV.Row, 8) = rs.Fields("Sales_Tax") & "%"
                    grdGRV.TextMatrix(grdGRV.Row, 10) = rs.Fields("Department_No")
                    grdGRV.TextMatrix(grdGRV.Row, 9) = 0
                End If
                rs.Close
                grdGRV.Col = 5
                frmMain.Toolbar1.Buttons(2).Enabled = True
                frmMain.Toolbar1.Buttons(3).Enabled = False
                frmMain.Toolbar1.Buttons(4).Enabled = True
            End If
        Case 46
            grdGRV.RemoveItem (grdGRV.Row)
            Calculate_Totals
            If .Rows = 1 Then
                txtInvoice.SetFocus
                frmMain.Toolbar1.Buttons(2).Enabled = False
                frmMain.Toolbar1.Buttons(3).Enabled = True
                frmMain.Toolbar1.Buttons(4).Enabled = False
            End If
        Case 45, 48 To 57, 96 To 105, 109, 110, 189
            Select Case grdGRV.Col
                Case 1, 2, 8
                Case Else
                    .EditCell
            End Select
        Case 37 'left
            If .Col = 1 Then
                KeyCode = 0
                .Col = 9
            End If
        Case 38 'up
            If grdGRV.Row = 1 Then
                txtInvoice.SetFocus
            End If
            If Trim(grdGRV.TextMatrix(grdGRV.Row, 1)) = "" Then
                grdGRV.RemoveItem (grdGRV.Row)
            End If
        Case 39 'Right
            If .Col = 9 Then
                KeyCode = 0
                .Col = 1
            End If
        Case 40 'down
            If Trim(grdGRV.TextMatrix(grdGRV.Row, 1)) = "" Then
                KeyCode = 0
                txtSubtotal.SetFocus
                If grdGRV.Rows > 2 Then grdGRV.RemoveItem (grdGRV.Row)
            Else
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .TextMatrix(grdGRV.Row, 2) = .TextMatrix(grdGRV.Row - 1, 2)
                    grdGRV.ShowCell .Row, 0
                    .Col = 1
                End If
            End If
    End Select
End With
End Sub
Private Sub grdGRV_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(grdGRV.EditText, ".") <> 0 And KeyAscii = 46 Then
        KeyAscii = 0
    End If
    Select Case KeyAscii
        Case 8, 13, 27, 45, 46, 48 To 57
        Case Else
            If Col > 2 Then KeyAscii = 0
    End Select
End Sub
Private Sub grdGRV_LostFocus()
    cmdProducts.Visible = False
    cmdHistory.Visible = False
End Sub

Private Sub grdSupp_DblClick()
    cmdSupplier.Value = Up
    cmdSupplier_Click
    txtAddress.Tag = grdSupp.TextMatrix(grdSupp.Row, 0)
    cmbSuppliers.Text = grdSupp.TextMatrix(grdSupp.Row, 1) & " (" & grdSupp.TextMatrix(grdSupp.Row, 0) & ")"
    LoadInfo
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
    End If
    rs.Close
    txtInvoice.SetFocus
    On Error GoTo 0
End Sub

Private Sub txtAddress_GotFocus()
    cmdStrip.Enabled = False
End Sub
Private Sub txtDiscount_Change()
    Calculate_Totals
End Sub
Private Sub txtDiscount_Click()
    txtDiscount.SelStart = 0
    txtDiscount.SelLength = Len(txtDiscount.Text)
    cmdStrip.Enabled = False
End Sub
Private Sub txtDiscount_GotFocus()
    txtDiscount.SelStart = 0
    txtDiscount.SelLength = Len(txtDiscount.Text)
    cmdStrip.Enabled = False
End Sub
Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            grdGRV.SetFocus
            grdGRV.Col = 1
        Case 40
            txtUllages.SetFocus
    End Select
End Sub
Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If InStr(txtDiscount.Text, ".") <> 0 And KeyAscii = 46 Then
        If txtDiscount.SelLength <> Len(txtDiscount.Text) Then
            KeyAscii = 0
        End If
    End If
    Select Case KeyAscii
        Case 8, 45, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtDiscount_LostFocus()
    txtDiscount.Text = Format(txtDiscount, "0.00")
End Sub

Private Sub txtInvoice_GotFocus()
    cmdStrip.Enabled = False
End Sub
Private Sub txtInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If txtInvoice.Text = "" And cmbSuppliers.Text <> "" Then
                MsgBox "Please Enter an Invoice Number", vbCritical, "HeroPOS"
                txtInvoice.SetFocus
                KeyCode = 0
                Exit Sub
            End If
            ActiveReadServer "Select Invoice_no from Purchase_Journal where Invoice_No = '" & txtInvoice.Text & "' and Supplier_No = '" & txtAddress.Tag & "'"
            If rs.RecordCount > 0 Then
                MsgBox "This Invoice Number was already Entered", vbCritical, "HeroPOS"
                txtInvoice.SetFocus
                KeyCode = 0
                rs.Close
                Exit Sub
            End If
            rs.Close
            
            If grdGRV.Rows = 1 Then
                grdGRV.Rows = grdGRV.Rows + 1
                grdGRV.Row = 1
                grdGRV.Col = 2
            End If
            grdGRV.SetFocus
        Case 38
            cmbSuppliers.SetFocus
        Case 40
            If txtInvoice.Text = "" And cmbSuppliers.Text <> "" Then
                MsgBox "Please Enter an Invoice Number", vbCritical, "HeroPOS"
                txtInvoice.SetFocus
                KeyCode = 0
                Exit Sub
            End If
            ActiveReadServer "Select Invoice_no from Purchase_Journal where Invoice_No = '" & txtInvoice.Text & "' and Supplier_No = '" & txtAddress.Tag & "'"
            If rs.RecordCount > 0 Then
                MsgBox "This Invoice Number was already Entered", vbCritical, "HeroPOS"
                txtInvoice.SetFocus
                KeyCode = 0
                rs.Close
                Exit Sub
            End If
            rs.Close
            If grdGRV.Rows = 1 Then
                grdGRV.Rows = grdGRV.Rows + 1
                grdGRV.Row = 1
                grdGRV.Col = 2
            End If
            grdGRV.Row = 1
                grdGRV.Col = 1
            grdGRV.SetFocus
    End Select
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = Asc("`")
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 48 To 57
    End Select
End Sub
Private Sub txtInvoice_LostFocus()
    On Error Resume Next
    ActiveReadServer "Select Invoice_no from Purchase_Journal where Invoice_No = '" & txtInvoice.Text & "' and Supplier_No = '" & txtAddress.Tag & "'"
    If rs.RecordCount > 0 Then
        MsgBox "This Invoice Number was already Entered", vbCritical, "HeroPOS"
        txtInvoice.SetFocus
        KeyCode = 0
        rs.Close
        On Error GoTo 0
        Exit Sub
    End If
    rs.Close
    On Error GoTo 0
End Sub
Private Sub txtSubtotal_Click()
    txtSubtotal.SelStart = 0
    txtSubtotal.SelLength = Len(txtSubtotal.Text)
    cmdStrip.Enabled = False
End Sub
Private Sub txtSubTotal_GotFocus()
    txtSubtotal.SelStart = 0
    txtSubtotal.SelLength = Len(txtSubtotal.Text)
    cmdStrip.Enabled = False
    cmdSundries.Value = Up
    picSundries.Visible = False
    Label18 = "Calculated Totals"
End Sub

Private Sub txtSubTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            grdGRV.SetFocus
            grdGRV.Col = 1
        Case 40
            txtVat.SetFocus
    End Select
End Sub

Private Sub txtSubTotal_KeyPress(KeyAscii As Integer)
    If InStr(txtSubtotal.Text, ".") <> 0 And KeyAscii = 46 Then
        If txtSubtotal.SelLength <> Len(txtSubtotal.Text) Then
            KeyAscii = 0
        End If
    End If
    Select Case KeyAscii
        Case 8, 45, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtSubTotal_LostFocus()
    txtSubtotal.Text = Format(txtSubtotal, "0.00")
End Sub

Private Sub txtTotal_GotFocus()
    txtTotal.SelStart = 0
    txtTotal.SelLength = Len(txtTotal.Text)
    cmdStrip.Enabled = False
    cmdSundries.Value = Up
    picSundries.Visible = False
    Label18 = "Calculated Totals"
End Sub

Private Sub txtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtVat.SetFocus
        Case 13, 40
            cmbSuppliers.SetFocus
    End Select
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    If InStr(txtTotal.Text, ".") <> 0 And KeyAscii = 46 Then
        If txtTotal.SelLength <> Len(txtTotal.Text) Then
            KeyAscii = 0
        End If
    End If
    Select Case KeyAscii
        Case 8, 45, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtTotal_LostFocus()
    txtTotal.Text = Format(txtTotal, "0.00")
End Sub

Private Sub txtTransport_Change()
    Calculate_Totals
End Sub
Private Sub txtTransport_Click()
    txtTransport.SelStart = 0
    txtTransport.SelLength = Len(txtUllages.Text)
    cmdStrip.Enabled = False
End Sub
Private Sub txtTransport_GotFocus()
    txtTransport.SelStart = 0
    txtTransport.SelLength = Len(txtUllages.Text)
    cmdStrip.Enabled = False
End Sub
Private Sub txtTransport_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtDiscount.SetFocus
        Case 13, 40
            cmbSuppliers.SetFocus
    End Select
End Sub
Private Sub txtTransport_KeyPress(KeyAscii As Integer)
    If InStr(txtTransport.Text, ".") <> 0 And KeyAscii = 46 Then
        If txtTransport.SelLength <> Len(txtTransport.Text) Then
            KeyAscii = 0
        End If
    End If
    Select Case KeyAscii
        Case 8, 45, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtTransport_LostFocus()
    txtTransport.Text = Format(txtTransport, "0.00")
End Sub
Private Sub txtUllages_Change()
    Calculate_Totals
End Sub
Private Sub txtUllages_Click()
    txtUllages.SelStart = 0
    txtUllages.SelLength = Len(txtUllages.Text)
    cmdStrip.Enabled = False
End Sub
Private Sub txtUllages_GotFocus()
    txtUllages.SelStart = 0
    txtUllages.SelLength = Len(txtUllages.Text)
    cmdStrip.Enabled = False
End Sub

Private Sub txtUllages_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtDiscount.SetFocus
        Case 40
            txtTransport.SetFocus
    End Select
End Sub

Private Sub txtUllages_KeyPress(KeyAscii As Integer)
    If InStr(txtUllages.Text, ".") <> 0 And KeyAscii = 46 Then
        If txtUllages.SelLength <> Len(txtUllages.Text) Then
            KeyAscii = 0
        End If
    End If
    Select Case KeyAscii
        Case 8, 45, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtUllages_LostFocus()
    txtUllages.Text = Format(txtUllages, "0.00")
End Sub
Private Sub txtVat_GotFocus()
    txtVat.SelStart = 0
    txtVat.SelLength = Len(txtVat.Text)
    cmdStrip.Enabled = False
    cmdSundries.Value = Up
    picSundries.Visible = False
    Label18 = "Calculated Totals"
End Sub
Private Sub txtVAT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            txtSubtotal.SetFocus
        Case 40
            txtTotal.SetFocus
    End Select
End Sub

Private Sub txtVat_KeyPress(KeyAscii As Integer)
    If InStr(txtVat.Text, ".") <> 0 And KeyAscii = 46 Then
        If txtVat.SelLength <> Len(txtVat.Text) Then
            KeyAscii = 0
        End If
    End If
    Select Case KeyAscii
        Case 8, 45, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtVat_LostFocus()
    txtVat.Text = Format(txtVat, "0.00")
End Sub
