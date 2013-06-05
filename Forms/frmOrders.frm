VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmOrders 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   9630
   ClientLeft      =   600
   ClientTop       =   780
   ClientWidth     =   14205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleMode       =   0  'User
   ScaleWidth      =   43115.82
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid grdClaims 
      Height          =   4515
      Left            =   180
      TabIndex        =   0
      Top             =   3030
      Width           =   13965
      _cx             =   24633
      _cy             =   7964
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
      FormatString    =   $"frmOrders.frx":0000
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   11700
         Width           =   1005
      End
   End
   Begin MSForms.Image Image7 
      Height          =   1875
      Left            =   10740
      Top             =   7680
      Width           =   3405
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "6006;3307"
   End
   Begin MSForms.Image Image6 
      Height          =   1875
      Left            =   6060
      Top             =   7680
      Width           =   4575
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "8070;3307"
   End
   Begin MSForms.Image Image5 
      Height          =   1875
      Left            =   60
      Top             =   7680
      Width           =   5895
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "10398;3307"
   End
   Begin MSForms.Image Image4 
      Height          =   4605
      Left            =   60
      Top             =   3000
      Width           =   14085
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "24844;8123"
   End
   Begin MSForms.Image Image3 
      Height          =   555
      Left            =   60
      Top             =   2370
      Width           =   14085
      BackColor       =   16707305
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "24844;979"
   End
   Begin MSForms.Image Image2 
      Height          =   2235
      Left            =   6930
      Top             =   60
      Width           =   7215
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "12726;3942"
   End
   Begin MSForms.Image Image1 
      Height          =   2235
      Left            =   60
      Top             =   60
      Width           =   6765
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   2
      Size            =   "11933;3942"
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image3_Click()

End Sub
