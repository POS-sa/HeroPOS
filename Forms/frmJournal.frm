VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmJournal 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picClose 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   60
      ScaleHeight     =   1305
      ScaleWidth      =   5655
      TabIndex        =   31
      Top             =   6780
      Width           =   5655
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Select and Account to pass a Journal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   885
         Left            =   570
         TabIndex        =   32
         Top             =   150
         Width           =   4545
      End
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   555
      Left            =   4590
      TabIndex        =   0
      Top             =   7440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   979
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx cmdOK 
      Height          =   555
      Left            =   3390
      TabIndex        =   1
      Top             =   7440
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   979
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "&Ok"
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
   End
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   765
      Left            =   1530
      TabIndex        =   2
      Top             =   900
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1349
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmJournal.frx":0000
   End
   Begin MSComCtl2.DTPicker DTStart 
      Height          =   345
      Left            =   7890
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   450
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   73203715
      CurrentDate     =   38862
   End
   Begin MSComCtl2.DTPicker DTStop 
      Height          =   345
      Left            =   7890
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   870
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   73203715
      CurrentDate     =   38862
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   285
      Left            =   150
      TabIndex        =   22
      Top             =   180
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Account No..."
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
   Begin VSFlex8Ctl.VSFlexGrid grdJournal 
      Height          =   600
      Left            =   60
      TabIndex        =   23
      Top             =   6780
      Width           =   5625
      _cx             =   9922
      _cy             =   1058
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   1500
      ColWidthMax     =   1500
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmJournal.frx":0082
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   25
         Top             =   11700
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   24
         Top             =   5670
         Width           =   1005
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdList 
      Height          =   4860
      Left            =   90
      TabIndex        =   26
      Top             =   1860
      Width           =   10095
      _cx             =   17806
      _cy             =   8572
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
      FormatString    =   $"frmJournal.frx":00FA
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
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   28
         Top             =   11700
         Width           =   1005
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   27
         Top             =   5670
         Width           =   1005
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdAcc 
      Height          =   4860
      Left            =   90
      TabIndex        =   3
      Top             =   1860
      Width           =   10095
      _cx             =   17806
      _cy             =   8572
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
      FormatString    =   $"frmJournal.frx":0172
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   11700
         Width           =   1005
      End
   End
   Begin MSComCtl2.DTPicker DTDated 
      Height          =   345
      Left            =   1140
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7440
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   73203715
      CurrentDate     =   38862
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Journal Date:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   7500
      Width           =   1095
   End
   Begin MSForms.Image Image15 
      Height          =   315
      Left            =   7440
      Top             =   7590
      Width           =   2655
      BorderColor     =   12632256
      Size            =   "4683;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image14 
      Height          =   315
      Left            =   7440
      Top             =   7230
      Width           =   2655
      BorderColor     =   12632256
      Size            =   "4683;556"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image13 
      Height          =   315
      Left            =   7440
      Top             =   6870
      Width           =   2655
      BorderColor     =   12632256
      Size            =   "4683;556"
      VariousPropertyBits=   19
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5835
      TabIndex        =   21
      Top             =   6930
      Width           =   1545
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vat:"
      Height          =   195
      Left            =   5835
      TabIndex        =   20
      Top             =   7275
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   5835
      TabIndex        =   19
      Top             =   7650
      Width           =   1545
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
      Left            =   7680
      TabIndex        =   18
      Top             =   6870
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
      Left            =   7680
      TabIndex        =   17
      Top             =   7230
      Width           =   2295
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
      Left            =   7680
      TabIndex        =   16
      Top             =   7590
      Width           =   2295
   End
   Begin VB.Label lblDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7980
      TabIndex        =   15
      Top             =   150
      Width           =   1995
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Listed To:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   14
      Top             =   930
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
      Left            =   6840
      TabIndex        =   13
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Listed From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   525
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
   Begin VB.Label lblAcc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<Select an Account>"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      TabIndex        =   9
      Top             =   225
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   165
      Left            =   390
      TabIndex        =   8
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name: "
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   570
      Width           =   945
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      TabIndex        =   6
      Top             =   555
      Width           =   4275
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
   Begin MSForms.Image Image4 
      Height          =   285
      Left            =   1440
      Top             =   510
      Width           =   4515
      BorderColor     =   12632256
      BackColor       =   16051176
      Size            =   "7964;503"
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
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   7890
      Top             =   60
      Width           =   2235
      BorderColor     =   12632256
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "3942;609"
   End
   Begin MSForms.Image Image5 
      Height          =   1215
      Left            =   5760
      Top             =   6780
      Width           =   4425
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7805;2143"
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonEx1_Click()
    lblAcc.Caption = "<Select an Account>"
    lblName.Caption = ""
    txtAddress.Text = ""
    grdList.Visible = True
    Form_Activate
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    Dim i As Integer
    Select Case frmJournal.Tag
        Case "Creditor"
            Journal_No = 0
            ActiveReadServer1 "Select isnull(max(Invoice_No),0)+1 as Receipt_No from Supplier_Accounts where Transaction_Type = 'Journal'"
            Journal_No = rs1.Fields("Receipt_No")
            rs1.Close
            ActiveUpdateServer "INSERT INTO [Supplier_Accounts]([User_No],[Date_Time],[Transaction_Type], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Tender_Type],[Ref_No],[Payment_No])" & _
            "VALUES(" & UserRecord.User_Number & ",'" & DTDated.Value & "','Journal','" & Journal_No & "','" & lblAcc.Caption & "'," & grdJournal.TextMatrix(1, 0) & "," & grdJournal.TextMatrix(1, 1) & ",0,'','" & grdJournal.TextMatrix(1, 2) & "',0)"
            
            ActiveReadServer2 "Select * from Supplier_Accounts where Account_No = '" & lblAcc.Caption & "' order by Date_Time"
            Balance = 0
            While Not rs2.EOF
                Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
                ActiveUpdateServer "Update Supplier_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
                rs2.MoveNext
            Wend
            rs2.Close
            ActiveUpdateServer "Update Suppliers set Balance = " & Balance & " Where Supplier_no = '" & lblAcc.Caption & "'"
        
'********** Debtormess Invoice_No must come from Sales_journal not Debtor_accounts ********

'        If TillData.DocNo = 0 Then
'        ActiveReadServer "Select (Select isnull(Max(convert(int,Trans_No)),0)+1 from Sales_Journal) as Trans_No,(Select isnull(Max(convert(int,Invoice_No)),0)+1 from Sales_Journal) as Invoice_No"
'            If rs.RecordCount > 0 Then
'            TillData.DocNo = rs.Fields("Invoice_No")
'            TillData.TransNo = rs.Fields("Trans_No")
'            End If
'        rs.Close
'        DoEvents
'        ActiveUpdateServer "Insert into Sales_Journal (Date_Time,Invoice_No,Trans_No,Table_No,Tab_No,Function_Key,User_Overide)  values (Getdate()," & TillData.DocNo & "," & TillData.TransNo & ",'" & TillData.TableNo & "','" & TillData.TabNo & "',14," & UserRecord.User_Number & ")"
'    End If

        
        Case "Debtor"
            Journal_No = 0
            'Must not have VAT
            ActiveReadServer1 "Select isnull(max(Invoice_No),0)+1 as Receipt_No from Debtor_Accounts where Transaction_Type = 'Journal'"
            Journal_No = rs1.Fields("Receipt_No")
            rs1.Close
            
            ActiveUpdateServer "INSERT INTO [Debtor_Accounts]([User_No],[Date_Time],[Transaction_Type], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Tender_Type],[Ref_No],[Payment_No])" & _
            "VALUES(" & UserRecord.User_Number & ",'" & DTDated.Value & "','Journal','" & Journal_No & "','" & lblAcc.Caption & "'," & grdJournal.TextMatrix(1, 0) & "," & grdJournal.TextMatrix(1, 1) & ",0,'','" & grdJournal.TextMatrix(1, 2) & "',0)"
            
            ActiveReadServer2 "Select * from Debtor_Accounts where Account_No = '" & lblAcc.Caption & "' order by Date_Time"
            Balance = 0
            While Not rs2.EOF
                Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
                ActiveUpdateServer "Update Debtor_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
                rs2.MoveNext
            Wend
            rs2.Close
            ActiveUpdateServer "Update Debtors set Balance = " & Balance & " Where Debtor_no = '" & lblAcc.Caption & "'"
            If frmMain.cmdBar(2).Enabled = False Then
                For i = 0 To frmMain.cmdMenu.Count - 1
                    If frmMain.cmdMenu(i).FontTextCaption.Bold = True Then
                        Select Case i
                            Case 9: frmMain.cmdMenu_Click i
                            Case 10: frmMain.cmdMenu_Click i
                            Case 11: frmMain.cmdMenu_Click i
                            Case 12: frmMain.cmdMenu_Click i
                        End Select
                    End If
                Next i
            End If
    End Select
    Unload Me
End Sub
Private Sub DTStart_CloseUp()
    grdList_DblClick
End Sub
Private Sub DTStop_CloseUp()
    grdList_DblClick
End Sub
Private Sub Form_Activate()
    Select Case frmJournal.Tag
        Case "Debtor"
            frmJournal.Caption = " Pass a Debtors Journal"
            grdList.Cols = 5
            grdList.TextMatrix(0, 0) = " Debtor Number"
            grdList.TextMatrix(0, 1) = " Debtor Name"
            grdList.TextMatrix(0, 2) = "Contact Person"
            grdList.TextMatrix(0, 3) = "Tel.Number"
            grdList.TextMatrix(0, 4) = "Balance "
            grdList.ColWidth(0) = grdList.Width * 0.2
            grdList.ColWidth(1) = grdList.Width * 0.3
            grdList.ColWidth(2) = grdList.Width * 0.22
            grdList.ColWidth(3) = grdList.Width * 0.15
            grdList.ColWidth(4) = grdList.Width * 0.13
            grdList.ColAlignment(0) = flexAlignLeftCenter
            grdList.ColAlignment(1) = flexAlignLeftCenter
            grdList.ColAlignment(2) = flexAlignLeftCenter
            grdList.ColAlignment(3) = flexAlignLeftCenter
            grdList.ColAlignment(4) = flexAlignRightCenter
            ActiveReadServer "Select * from Debtors order by Debtor_No"
            grdList.Rows = rs.RecordCount + 1
            i = 0
            While Not rs.EOF
                i = i + 1
                grdList.TextMatrix(i, 0) = rs.Fields("Debtor_No")
                grdList.TextMatrix(i, 1) = rs.Fields("Debtor_Name")
                grdList.TextMatrix(i, 2) = rs.Fields("Contact_Person") & ""
                grdList.TextMatrix(i, 3) = rs.Fields("Business_Tel") & ""
                grdList.TextMatrix(i, 4) = Format(rs.Fields("Balance"), "0.00")
                rs.MoveNext
            Wend
            rs.Close
            If grdList.Rows > 1 Then grdList.Row = 1
        Case "Creditor"
            frmJournal.Caption = " Pass a Creditors Journal"
            grdList.Cols = 5
            grdList.TextMatrix(0, 0) = " Creditor Number"
            grdList.TextMatrix(0, 1) = " Creditor Name"
            grdList.TextMatrix(0, 2) = "Contact Person"
            grdList.TextMatrix(0, 3) = "Tel.Number"
            grdList.TextMatrix(0, 4) = "Balance "
            grdList.ColWidth(0) = grdList.Width * 0.2
            grdList.ColWidth(1) = grdList.Width * 0.3
            grdList.ColWidth(2) = grdList.Width * 0.22
            grdList.ColWidth(3) = grdList.Width * 0.15
            grdList.ColWidth(4) = grdList.Width * 0.13
            grdList.ColAlignment(0) = flexAlignLeftCenter
            grdList.ColAlignment(1) = flexAlignLeftCenter
            grdList.ColAlignment(2) = flexAlignLeftCenter
            grdList.ColAlignment(3) = flexAlignLeftCenter
            grdList.ColAlignment(4) = flexAlignRightCenter
            grdList.Rows = 1
            ActiveReadServer "Select * from Suppliers order by Supplier_Name"
            grdList.Rows = rs.RecordCount + 1
            i = 0
            While Not rs.EOF
                i = i + 1
                grdList.TextMatrix(i, 0) = rs.Fields("Supplier_No")
                grdList.TextMatrix(i, 1) = rs.Fields("Supplier_Name")
                grdList.TextMatrix(i, 2) = rs.Fields("Contact_Person")
                grdList.TextMatrix(i, 3) = rs.Fields("Business_Tel")
                grdList.TextMatrix(i, 4) = Format(rs.Fields("Balance"), "0.00")
                rs.MoveNext
            Wend
            rs.Close
            If grdList.Rows > 1 Then grdList.Row = 1
    End Select
    Screen.MousePointer = 11
    DoEvents
    If grdAcc.Tag = "1" Then
        grdAcc.Tag = ""
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
    grdAcc.Rows = 1
    grdAcc.Cols = 7
    grdAcc.TextMatrix(0, 0) = "Dated"
    grdAcc.TextMatrix(0, 1) = "Transaction Details"
    grdAcc.TextMatrix(0, 2) = "Debit"
    grdAcc.TextMatrix(0, 3) = "Credit"
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
    DTDated.Value = Date
    grdJournal.TextMatrix(0, 0) = "Debit"
    grdJournal.TextMatrix(0, 1) = "Credit"
    grdJournal.TextMatrix(0, 2) = "Reason"
    grdJournal.TextMatrix(1, 0) = "0.00"
    grdJournal.TextMatrix(1, 1) = "0.00"
    grdJournal.TextMatrix(1, 2) = "<None>"
End Sub

Private Sub grdJournal_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
        Case 0
            If grdJournal.ValueMatrix(1, 0) <> 0 Then
                grdJournal.TextMatrix(1, 1) = "0.00"
            End If
            grdJournal.TextMatrix(1, 0) = Format(grdJournal.ValueMatrix(1, 0), "0.00")
        Case 1
            If grdJournal.ValueMatrix(1, 1) <> 0 Then
                grdJournal.TextMatrix(1, 0) = "0.00"
            End If
            grdJournal.TextMatrix(1, 1) = Format(grdJournal.ValueMatrix(1, 1), "0.00")
        Case 2
            If Trim(grdJournal.TextMatrix(1, 2)) = "" Then
                grdJournal.TextMatrix(1, 2) = "<None>"
            End If
            grdJournal.TextMatrix(1, 2) = UCase(Left(grdJournal.TextMatrix(1, 2), 1)) & Mid(grdJournal.TextMatrix(1, 2), 2)
    End Select
    cmdOk.Enabled = False
    If grdJournal.ValueMatrix(1, 0) + grdJournal.ValueMatrix(1, 1) <> 0 Then
        If grdJournal.TextMatrix(1, 2) <> "<None>" Then
            If lblAcc.Caption <> "<Select an Account>" Then
                cmdOk.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub grdJournal_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If grdJournal.ValueMatrix(1, 0) + grdJournal.ValueMatrix(1, 1) <> 0 Then
        If grdJournal.TextMatrix(1, 2) <> "<None>" Then
            If lblAcc.Caption <> "<Select an Account>" Then
                cmdOk.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub grdJournal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 45, 48 To 57
            Select Case grdList.Col
                Case 0, 1, 2
                    grdJournal.EditCell
            End Select
        Case 96 To 105, 109, 110, 189
            Select Case grdList.Col
                Case 2
                    grdJournal.EditCell
            End Select
    End Select
End Sub
Private Sub grdJournal_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case grdJournal.Col
        Case 0, 1
            If InStr(grdJournal.EditText, ".") <> 0 And KeyAscii = 46 Then
                KeyAscii = 0
            End If
            Select Case KeyAscii
                Case 8, 13, 27, 45, 46, 48 To 57
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub
Private Sub grdList_DblClick()
    lblAcc.Caption = grdList.TextMatrix(grdList.Row, 0)
    picClose.Visible = False
    grdList.Visible = False
    lblSub.Caption = "0.00"
    If frmJournal.Tag = "Debtor" Then
        ActiveReadServer "Select * from Debtors where Debtor_No = '" & lblAcc.Caption & "'"
        If rs.RecordCount > 0 Then
            lblName.Caption = rs.Fields("Debtor_Name")
            txtAddress.Text = rs.Fields("Address")
            lblDate.Caption = Format(Date, "ddd dd MMM yyyy")
        End If
        rs.Close
    End If
    If frmJournal.Tag = "Creditor" Then
        ActiveReadServer "Select * from Suppliers where Supplier_No = '" & lblAcc.Caption & "'"
        If rs.RecordCount > 0 Then
            lblAcc.Caption = lblAcc.Caption
            lblName.Caption = rs.Fields("Supplier_Name")
            txtAddress.Text = rs.Fields("Address")
            lblDate.Caption = Format(Date, "ddd dd MMM yyyy")
        End If
        rs.Close
    End If
    Select Case frmJournal.Tag
        Case "Creditor"
            ActiveReadServer "Select 1 as Line_No,'" & DTStart.Value & "' as Date_Time,'' as Invoice_No,'' as Payment_No,'Opening Balance' as Transaction_Type" & _
            " ,Account_No,Sum(Debit)as Debit,Sum(Credit)as Credit,Sum(Debit)-Sum(Credit) as Balance," & _
            " 0 as User_No,'' as Tender_Type,'' as Ref_No" & _
            " from Supplier_Accounts where Account_No = '" & lblAcc.Caption & "'" & _
            " and (Date_Time < '" & DTStart.Value & "' )" & _
            " group by Account_No " & _
            " Union" & _
            " Select * from Supplier_Accounts where Account_No = '" & lblAcc.Caption & "'" & _
            " and (Date_Time > '" & DTStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & DTStop.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "') order by Date_Time"
        Case "Debtor"
            ActiveReadServer "Select 1 as Line_No,'" & DTStart.Value & "' as Date_Time,0 as Invoice_No,'Opening Balance' as Transaction_Type" & _
            " ,Account_No,Sum(Debit)as Debit,Sum(Credit)as Credit,Sum(Debit)-Sum(Credit) as Balance," & _
            " 0 as User_No,'' as Tender_Type,'' as Ref_No,'' as Payment_No" & _
            " from Debtor_Accounts where Account_No = '" & lblAcc.Caption & "'" & _
            " and (Date_Time < '" & DTStart.Value & "' )" & _
            " group by Account_No" & _
            " Union" & _
            " Select * from Debtor_Accounts where Account_No = '" & lblAcc.Caption & "'" & _
            " and (Date_Time > '" & DTStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & DTStop.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "') order by Date_Time"
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
        lblSub.Caption = Format(Balance - Totvat, "0.00")
    End If
    lblVat.Caption = Format(Totvat, "0.00")
    If grdAcc.Rows > 1 Then grdAcc.Row = 1
    DoEvents
    Screen.MousePointer = 0
    
    
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            grdList_DblClick
    End Select
End Sub
