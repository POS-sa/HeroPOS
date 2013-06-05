VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRooms 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   15105
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid grdRooms 
      Height          =   1920
      Left            =   0
      TabIndex        =   26
      Top             =   7410
      Width           =   13695
      _cx             =   24156
      _cy             =   3387
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRooms.frx":0000
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
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   27
         Top             =   11700
         Width           =   1005
      End
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
      TabIndex        =   23
      Top             =   6660
      Width           =   14865
      Begin btButtonEx.ButtonEx cmdUp 
         Height          =   465
         Left            =   13110
         TabIndex        =   24
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
      Begin VB.Label lblRoomDet 
         Caption         =   "Room Details..."
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
         Left            =   270
         TabIndex        =   25
         Top             =   240
         Width           =   3135
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   2
         Left            =   3870
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   1
         Left            =   3540
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image4 
         Height          =   90
         Index           =   0
         Left            =   3210
         Top             =   570
         Width           =   285
         BackColor       =   16761024
         Size            =   "503;159"
      End
      Begin MSForms.Image Image6 
         Height          =   90
         Left            =   150
         Top             =   570
         Width           =   3015
         BackColor       =   16761024
         Size            =   "5318;159"
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
      Begin MSForms.Image picSeperate1 
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   13755
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "24262;185"
      End
   End
   Begin VB.PictureBox picLocation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   8640
      ScaleHeight     =   3195
      ScaleWidth      =   4335
      TabIndex        =   45
      Top             =   2790
      Visible         =   0   'False
      Width           =   4365
      Begin VSFlex8Ctl.VSFlexGrid grdLocations 
         Height          =   3255
         Left            =   -30
         TabIndex        =   46
         Top             =   -30
         Width           =   4395
         _cx             =   7752
         _cy             =   5741
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         BackColorSel    =   15391677
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRooms.frx":0078
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
      End
   End
   Begin VB.TextBox txtShower1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   43
      Top             =   6270
      Width           =   945
   End
   Begin VB.TextBox txtSpa 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   42
      Top             =   5910
      Width           =   945
   End
   Begin VB.TextBox txtShowerb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   41
      Top             =   5550
      Width           =   945
   End
   Begin VB.TextBox txtRoll 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   40
      Top             =   5190
      Width           =   945
   End
   Begin VB.TextBox txtCorner 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   39
      Top             =   4830
      Width           =   945
   End
   Begin VB.TextBox txtStraightbath 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   38
      Top             =   4470
      Width           =   945
   End
   Begin VB.TextBox txtBaby 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   37
      Top             =   3660
      Width           =   945
   End
   Begin VB.TextBox txtSingle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   36
      Top             =   3300
      Width           =   945
   End
   Begin VB.TextBox txtDouble 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   35
      Top             =   2940
      Width           =   945
   End
   Begin VB.TextBox txtQueen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   34
      Top             =   2580
      Width           =   945
   End
   Begin VB.TextBox txtKing 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      TabIndex        =   33
      Top             =   2220
      Width           =   945
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1680
      TabIndex        =   1
      Top             =   1530
      Width           =   4395
   End
   Begin VB.TextBox cmbRoom_No 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1680
      TabIndex        =   0
      Top             =   1140
      Width           =   1365
   End
   Begin VB.PictureBox picRoom 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   2730
      ScaleHeight     =   4485
      ScaleWidth      =   3345
      TabIndex        =   31
      Top             =   2070
      Width           =   3375
      Begin MSForms.Label Label4 
         Height          =   885
         Left            =   600
         TabIndex        =   32
         Top             =   1560
         Width           =   2175
         ForeColor       =   16777215
         BackColor       =   -2147483643
         VariousPropertyBits=   8388627
         Caption         =   "Click to add Picture"
         Size            =   "3836;1561"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Image picRooms 
         Height          =   4545
         Left            =   -30
         Top             =   -30
         Width           =   3465
         SizeMode        =   1
         Size            =   "6112;8017"
      End
   End
   Begin btButtonEx.ButtonEx cmdRates 
      Height          =   315
      Left            =   6390
      TabIndex        =   11
      Top             =   6060
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Rates..."
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
   Begin btButtonEx.ButtonEx cmdKing 
      Height          =   285
      Left            =   210
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2190
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "King Size"
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
   Begin btButtonEx.ButtonEx cmdSingle 
      Height          =   285
      Left            =   210
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3270
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Single"
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
   Begin btButtonEx.ButtonEx cmdDouble 
      Height          =   285
      Left            =   210
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2910
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Double"
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
   Begin btButtonEx.ButtonEx cmdQueen 
      Height          =   285
      Left            =   210
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Queen Size"
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
   Begin btButtonEx.ButtonEx cmdBaby 
      Height          =   285
      Left            =   210
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3630
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Baby Cot"
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
   Begin btButtonEx.ButtonEx cmdSpa 
      Height          =   285
      Left            =   210
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Spa Bath"
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
   Begin btButtonEx.ButtonEx cmdShowerb 
      Height          =   285
      Left            =   210
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Shower Bath"
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
   Begin btButtonEx.ButtonEx cmdRoll 
      Height          =   285
      Left            =   210
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Roll Top Bath"
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
   Begin btButtonEx.ButtonEx cmdCorner 
      Height          =   285
      Left            =   210
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Corner Bath"
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
   Begin btButtonEx.ButtonEx cmdStraight 
      Height          =   285
      Left            =   210
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Straight Bath"
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
   Begin btButtonEx.ButtonEx cmdShower 
      Height          =   285
      Left            =   210
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Shower"
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
   Begin VSFlex8Ctl.VSFlexGrid grdAmen 
      Height          =   4245
      Left            =   6420
      TabIndex        =   28
      Top             =   1170
      Width           =   6585
      _cx             =   11615
      _cy             =   7488
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
      BackColorSel    =   15391677
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRooms.frx":00F0
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
   End
   Begin MSComDlg.CommonDialog cmLogo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Logo"
      Filter          =   "*.jpg,*.bmp"
      InitDir         =   "c:\res\Rooms"
   End
   Begin btButtonEx.ButtonEx cmdLocation 
      Height          =   315
      Left            =   10590
      TabIndex        =   44
      Top             =   6060
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "Location Link...."
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
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   8
      Left            =   1500
      Top             =   6240
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   7
      Left            =   1500
      Top             =   5880
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   6
      Left            =   1500
      Top             =   5520
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   5
      Left            =   1500
      Top             =   5160
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   4
      Left            =   1500
      Top             =   4800
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   3
      Left            =   1500
      Top             =   4440
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   2
      Left            =   1500
      Top             =   3630
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   1
      Left            =   1500
      Top             =   3270
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image9 
      Height          =   285
      Left            =   1500
      Top             =   2910
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image8 
      Height          =   285
      Left            =   1500
      Top             =   2550
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Index           =   0
      Left            =   1500
      Top             =   2190
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;503"
   End
   Begin MSForms.Image Image3 
      Height          =   315
      Left            =   1530
      Top             =   1470
      Width           =   4575
      BackColor       =   16777215
      Size            =   "8070;556"
   End
   Begin MSForms.Image Image2 
      Height          =   315
      Index           =   0
      Left            =   1530
      Top             =   1080
      Width           =   1545
      BackColor       =   16777215
      Size            =   "2725;556"
   End
   Begin MSForms.Label Label6 
      Height          =   885
      Left            =   3360
      TabIndex        =   30
      Top             =   5040
      Width           =   2115
      ForeColor       =   16777215
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Click to add Picture"
      Size            =   "3731;1561"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   885
      Left            =   3360
      TabIndex        =   29
      Top             =   2640
      Width           =   2145
      ForeColor       =   16777215
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Click to add Picture"
      Size            =   "3784;1561"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image Picture1 
      Height          =   2265
      Left            =   2730
      Top             =   4290
      Width           =   3375
      SizeMode        =   1
      Size            =   "5953;3995"
   End
   Begin MSForms.Image Picture2 
      Height          =   1965
      Left            =   2730
      Top             =   2100
      Width           =   3375
      SizeMode        =   1
      Size            =   "5953;3466"
   End
   Begin MSForms.ComboBox cmbRates 
      Height          =   315
      Left            =   7860
      TabIndex        =   2
      Tag             =   "Up"
      Top             =   6060
      Width           =   2655
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4683;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line7 
      X1              =   6960
      X2              =   12990
      Y1              =   5670
      Y2              =   5670
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   5700
      TabIndex        =   10
      Top             =   5610
      Width           =   1185
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Rates"
      Size            =   "2090;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   6180
      TabIndex        =   9
      Top             =   780
      Width           =   1185
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Amenities"
      Size            =   "2090;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Line Line4 
      X1              =   7380
      X2              =   12990
      Y1              =   840
      Y2              =   840
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1185
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Bathrooms"
      Size            =   "2090;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Line Line3 
      X1              =   1470
      X2              =   2610
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   6090
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      X1              =   1230
      X2              =   6090
      Y1              =   840
      Y2              =   840
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   -270
      TabIndex        =   7
      Top             =   1890
      Width           =   1185
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Beds"
      Size            =   "2090;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Room Description:"
      Height          =   195
      Left            =   -180
      TabIndex        =   6
      Top             =   1530
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Room Number:"
      Height          =   195
      Left            =   390
      TabIndex        =   5
      Top             =   1170
      Width           =   1065
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   9
      Left            =   0
      TabIndex        =   4
      Top             =   780
      Width           =   1185
      ForeColor       =   12582912
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "General"
      Size            =   "2090;397"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   13
      Left            =   150
      TabIndex        =   3
      Top             =   300
      Width           =   3105
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Room Details"
      Size            =   "5477;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image Image1 
      Height          =   345
      Left            =   210
      Top             =   240
      Width           =   3195
      BackColor       =   16777215
      Size            =   "5636;609"
      VariousPropertyBits=   19
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbRates_Change()
    frmMain.Toolbar1.Buttons(4).Enabled = True
End Sub
Private Sub cmbRates_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub

Private Sub cmbRates_GotFocus()
    cmbRates.SelStart = 0
    cmbRates.SelLength = Len(cmbRates.Text)
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub cmbRates_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtShower1.SetFocus
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

Private Sub cmbRoom1_No_Change()
    If cmbRoom_No.Text <> "" And txtDescription.Text <> "" Then
         frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
    If grdRooms.FindRow(cmbRoom_No.Text, 0, 0, 0, 1) > 0 Then
        grdRooms.Row = grdRooms.FindRow(cmbRoom_No.Text, 0, 0, 0, 1)
        grdRooms.ShowCell grdRooms.Row, 0
    End If
End Sub
Private Sub cmbRoom_No_GotFocus()
    cmbRoom_No.SelStart = 0
    cmbRoom_No.SelLength = Len(cmbRoom_No.Text)
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub
Private Sub cmbRoom_No_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case 65, 66, 67
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cmdBaby_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtBaby.SetFocus
    Picture1.Tag = 5
    ActiveReadServer "Select Pic5 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture2.Picture = LoadPicture(Trim(rs.Fields("Pic5") & ""))
        If Trim(rs.Fields("Pic5") & "") <> "" Then
            Label5.Visible = False
        Else
            Label5.Visible = True
        End If
    Else
        Set Picture2.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdBaby_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdBaby_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdCorner_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtCorner.SetFocus
    Picture1.Tag = 7
    ActiveReadServer "Select Pic7 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture1.Picture = LoadPicture(Trim(rs.Fields("Pic7") & ""))
        If Trim(rs.Fields("Pic7") & "") <> "" Then
            Label6.Visible = False
        Else
            Label6.Visible = True
        End If
    Else
        Set Picture1.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdCorner_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdCorner_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdDouble_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtDouble.SetFocus
    Picture1.Tag = 3
    ActiveReadServer "Select Pic3 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture2.Picture = LoadPicture(Trim(rs.Fields("Pic3") & ""))
        If Trim(rs.Fields("Pic3") & "") <> "" Then
            Label5.Visible = False
        Else
            Label5.Visible = True
        End If
    Else
        Set Picture2.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdDouble_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdDouble_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdKing_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtKing.SetFocus
    Picture1.Tag = 1
    ActiveReadServer "Select Pic1 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture2.Picture = LoadPicture(Trim(rs.Fields("Pic1") & ""))
        If Trim(rs.Fields("Pic1") & "") <> "" Then
            Label5.Visible = False
        Else
            Label5.Visible = True
        End If
    Else
        Set Picture2.Picture = LoadPicture("")
    End If
    rs.Close
End Sub
Private Sub cmdKing_GotFocus()
    picRoom.Visible = False
End Sub
Private Sub cmdKing_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdLocation_Click()
    Select Case cmdLocation.Value
        Case 0
            picLocation.Visible = True
        Case 1
            picLocation.Visible = False
    End Select
End Sub

Private Sub cmdQueen_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtQueen.SetFocus
    Picture1.Tag = 2
    ActiveReadServer "Select Pic2 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture2.Picture = LoadPicture(Trim(rs.Fields("Pic2") & ""))
        If Trim(rs.Fields("Pic2") & "") <> "" Then
            Label5.Visible = False
        Else
            Label5.Visible = True
        End If
    Else
        Set Picture2.Picture = LoadPicture("")
    End If
    rs.Close
End Sub
Private Sub cmdQueen_GotFocus()
    picLocation.Visible = False
    picRoom.Visible = False
End Sub
Private Sub cmdQueen_LostFocus()
    picRoom.Visible = True
End Sub
Private Sub cmdRates_Click()
    Load frmRates
    frmRates.Tag = ""
    frmRates.Show vbModal
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub
Private Sub cmdRoll_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtRoll.SetFocus
    Picture1.Tag = 8
    ActiveReadServer "Select Pic8 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture1.Picture = LoadPicture(Trim(rs.Fields("Pic8") & ""))
        If Trim(rs.Fields("Pic8") & "") <> "" Then
            Label6.Visible = False
        Else
            Label6.Visible = True
        End If
    Else
        Set Picture1.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdRoll_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdRoll_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdShower_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtShower1.SetFocus
    Picture1.Tag = 11
    ActiveReadServer "Select Pic11 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture1.Picture = LoadPicture(Trim(rs.Fields("Pic11") & ""))
        If Trim(rs.Fields("Pic11") & "") <> "" Then
            Label6.Visible = False
        Else
            Label6.Visible = True
        End If
    Else
        Set Picture1.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdShower_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdShower_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdShowerb_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtShowerb.SetFocus
    Picture1.Tag = 9
    ActiveReadServer "Select Pic9 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture1.Picture = LoadPicture(Trim(rs.Fields("Pic9") & ""))
        If Trim(rs.Fields("Pic9") & "") <> "" Then
            Label6.Visible = False
        Else
            Label6.Visible = True
        End If
    Else
        Set Picture1.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdShowerb_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdShowerb_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdSingle_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtSingle.SetFocus
    Picture1.Tag = 4
    ActiveReadServer "Select Pic4 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture2.Picture = LoadPicture(Trim(rs.Fields("Pic4") & ""))
        If Trim(rs.Fields("Pic4") & "") <> "" Then
            Label5.Visible = False
        Else
            Label5.Visible = True
        End If
    Else
        Set Picture2.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdSingle_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdSingle_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdSpa_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtSpa.SetFocus
    Picture1.Tag = 10
    ActiveReadServer "Select Pic10 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture1.Picture = LoadPicture(Trim(rs.Fields("Pic10") & ""))
        If Trim(rs.Fields("Pic10") & "") <> "" Then
            Label6.Visible = False
        Else
            Label6.Visible = True
        End If
    Else
        Set Picture1.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdSpa_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdSpa_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdStraight_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    txtStraightbath.SetFocus
    Picture1.Tag = 6
    ActiveReadServer "Select Pic6 from Rooms where Room_No=" & Val(cmbRoom_No.Text)
    If rs.RecordCount > 0 Then
        Set Picture1.Picture = LoadPicture(Trim(rs.Fields("Pic6") & ""))
        If Trim(rs.Fields("Pic6") & "") <> "" Then
            Label6.Visible = False
        Else
            Label6.Visible = True
        End If
    Else
        Set Picture1.Picture = LoadPicture("")
    End If
    rs.Close
End Sub

Private Sub cmdStraight_GotFocus()
    picRoom.Visible = False
End Sub

Private Sub cmdStraight_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub cmdUp_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    Select Case cmdUp.Caption
        Case "5"
            grdRooms.SetFocus
            picHead.top = 0
            grdRooms.top = picHead.top + picHead.Height
            grdRooms.Height = frmRooms.Height - picHead.Height - 120
            DoEvents
            cmdUp.Caption = 6
        Case "6"
            cmdUp.Caption = "5"
            picHead.top = 6660
            grdRooms.top = 7410
            grdRooms.Height = 1920
            DoEvents
            cmbRoom_No.SetFocus
    End Select
End Sub

Private Sub Form_Activate()
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Rooms"
    If grdAmen.Rows = 1 Then grdAmen.Rows = 2
    cmbRoom_No.SetFocus
End Sub

Private Sub Form_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            KeyCode = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
            frmSplash.Show
            frmMain.Hide
    End Select
End Sub
Public Sub CreateRooms()
    cmbRoom_No.Text = ""
    txtDescription.Text = ""
    txtKing.Text = "0"
    txtQueen.Text = "0"
    txtDouble.Text = "0"
    txtSingle.Text = "0"
    txtBaby.Text = "0"
    txtStraightbath.Text = "0"
    txtCorner.Text = "0"
    txtRoll.Text = "0"
    txtShowerb.Text = "0"
    txtSpa.Text = "0"
    txtShower1.Text = "0"
    cmbRoom_No.SetFocus
    cmbRates.Text = "<Select a Rate>"
    frmMain.Toolbar1.Buttons(2).Enabled = False
End Sub
Public Sub SaveRooms()
    picLocation.Visible = False
    cmdLocation.Value = Up
    Amenstring = ""
    For i = 1 To grdAmen.Rows - 1
        If grdAmen.ValueMatrix(i, 1) = True Then
            Amenstring = Amenstring & i & "|"
        End If
    Next i
    If cmbRates.Text = "<Select a Rate>" Then
        Roomrate = 0
    Else
        Roomrate = Val(Mid(cmbRates.Text, 1, InStr(cmbRates.Text, "-") - 1))
    End If
    RoomLocation = Location_No
    For i = 1 To grdLocations.Rows - 1
        If grdLocations.ValueMatrix(i, 2) <> 0 Then
            RoomLocation = grdLocations.ValueMatrix(i, 0)
        End If
    Next i
    If Right(Amenstring, 1) = "|" Then Amenstring = Mid(Amenstring, 1, Len(Amenstring) - 1)
    Beds = Val(txtKing.Text) + Val(txtQueen.Text) + Val(txtDouble.Text) + Val(txtSingle.Text) + Val(txtBaby.Text)
    Baths = Val(txtStraightbath.Text) + Val(txtCorner.Text) + Val(txtRoll.Text) + Val(txtShowerb.Text) + Val(txtSpa.Text) + Val(txtShower1.Text)
    ActiveReadServer "Select * from Rooms where Room_No = '" & cmbRoom_No.Text & "'"
    If rs.RecordCount = 0 Then
        ActiveUpdateServer "INSERT INTO Rooms (Room_No,Description,King_Size,Queen_Size,Double_Bed,Single_Bed,Baby_Bed,Straight,Corner,Roll_Top,Shower_Bath,Spa,Shower,Bed_No,Bath_No, Ameneties,Room_Rate,Location_No)" & _
        " VALUES('" & cmbRoom_No.Text & "','" & txtDescription.Text & "','" & txtKing.Text & "','" & txtQueen.Text & "','" & txtDouble.Text & "','" & txtSingle.Text & "','" & txtBaby.Text & "','" & txtStraightbath.Text & "','" & txtCorner.Text & "','" & txtRoll.Text & "','" & txtShowerb.Text & "','" & txtSpa.Text & "','" & txtShower1.Text & "'," & Beds & "," & Baths & ",'" & Amenstring & "'," & Roomrate & ",'" & RoomLocation & "')"
        grdRooms.Rows = grdRooms.Rows + 1
        grdRooms.Tag = "1"
        grdRooms.Row = grdRooms.Rows - 1
        grdRooms.ShowCell grdRooms.Row, 0
        grdRooms.Tag = ""
    Else
        ActiveUpdateServer "UPDATE Rooms" & _
        " SET " & _
        " Room_No='" & cmbRoom_No.Text & "'," & _
        " Description='" & txtDescription.Text & "'," & _
        " King_Size=" & txtKing.Text & "," & _
        " Queen_Size='" & txtQueen.Text & "'," & _
        " Double_Bed='" & txtDouble.Text & "'," & _
        " Single_Bed='" & txtSingle.Text & "'," & _
        " Baby_Bed='" & txtBaby.Text & "'," & _
        " Straight='" & txtStraightbath.Text & "'," & _
        " Corner='" & txtCorner.Text & "'," & _
        " Roll_Top ='" & txtRoll.Text & "'," & _
        " Shower_Bath ='" & txtShowerb.Text & "'," & _
        " Spa='" & txtSpa.Text & "'," & _
        " Shower='" & txtShower1.Text & "'," & _
        " Bed_No ='" & Beds & "'," & _
        " Bath_No='" & Baths & "'," & _
        " Ameneties='" & Amenstring & "'," & _
        " Location_No='" & RoomLocation & "'," & _
        " Room_Rate='" & Roomrate & "'" & _
        " Where Room_No = '" & cmbRoom_No.Text & "'"
    End If
    rs.Close
    grdRooms.TextMatrix(grdRooms.Row, 0) = cmbRoom_No.Text
    grdRooms.TextMatrix(grdRooms.Row, 1) = txtDescription.Text
    grdRooms.TextMatrix(grdRooms.Row, 2) = Beds
    grdRooms.TextMatrix(grdRooms.Row, 3) = Baths
    grdRooms.TextMatrix(grdRooms.Row, 4) = cmbRates.Text
    MsgBox "Room update successful", vbInformation, "HeroPOS"
    frmMain.Toolbar1.Buttons(4).Enabled = False
    frmMain.Toolbar1.Buttons(2).Enabled = True
    frmMain.Toolbar1.Buttons(5).Enabled = True
End Sub
Private Sub Form_Load()
    On Error GoTo trap
    ChDir App.Path & "\Rooms"
    ChDir App.Path
    cmbRates.Clear
    cmbRates.AddItem "<Select a Rate>"
    ActiveReadServer "Select * from Room_Rates where Active=1 order by Rate_Type"
    While Not rs.EOF
        i = i + 1
        cmbRates.AddItem rs.Fields("Rate_Type") & " - " & rs.Fields("Description") & " > " & Format(rs.Fields("Room_Rate"), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdLocations.Cols = 3
    grdLocations.TextMatrix(0, 0) = "No"
    grdLocations.TextMatrix(0, 1) = "Location Name"
    grdLocations.TextMatrix(0, 2) = "Linked"
    grdLocations.ColAlignment(0) = flexAlignLeftCenter
    grdLocations.ColAlignment(1) = flexAlignLeftCenter
    grdLocations.ColAlignment(2) = flexAlignCentreCenter
    grdLocations.ColWidth(0) = grdLocations.Width * 0.1
    grdLocations.ColWidth(1) = grdLocations.Width * 0.65
    grdLocations.ColWidth(2) = grdLocations.Width * 0.25
    grdLocations.ColDataType(2) = flexDTBoolean
    grdLocations.Rows = 1
    ActiveReadServer "Select * from Locations where Loc_Type = 0 order by Location_No"
    While Not rs.EOF
        grdLocations.Rows = grdLocations.Rows + 1
        grdLocations.Row = grdLocations.Rows - 1
        grdLocations.TextMatrix(grdLocations.Row, 0) = rs.Fields("Location_No")
        grdLocations.TextMatrix(grdLocations.Row, 1) = rs.Fields("Loc_Name")
        grdLocations.TextMatrix(grdLocations.Row, 2) = ""
        rs.MoveNext
    Wend
    rs.Close
    cmbRates.Text = "<Select a Rate>"
    grdAmen.Rows = 18
    grdAmen.Cols = 2
    grdAmen.TextMatrix(0, 0) = "Description"
    grdAmen.TextMatrix(0, 1) = "Included"
    grdAmen.ColWidth(0) = 5000
    grdAmen.ColWidth(1) = 1500
    grdAmen.ColDataType(1) = flexDTBoolean
    grdAmen.TextMatrix(1, 0) = "Alarm Clock"
    grdAmen.TextMatrix(2, 0) = "Air Conditioning"
    grdAmen.TextMatrix(3, 0) = "Central Heating"
    grdAmen.TextMatrix(4, 0) = "Coffee Maker"
    grdAmen.TextMatrix(5, 0) = "Electric Towel Warmer"
    grdAmen.TextMatrix(6, 0) = "Electric Heater"
    grdAmen.TextMatrix(7, 0) = "Fire Place"
    grdAmen.TextMatrix(8, 0) = "Hairdryer"
    grdAmen.TextMatrix(9, 0) = "Internet Access"
    grdAmen.TextMatrix(10, 0) = "Kettle"
    grdAmen.TextMatrix(11, 0) = "Mini Bar"
    grdAmen.TextMatrix(12, 0) = "Radio"
    grdAmen.TextMatrix(13, 0) = "Satellite Television"
    grdAmen.TextMatrix(14, 0) = "Self Catering"
    grdAmen.TextMatrix(15, 0) = "Television"
    grdAmen.TextMatrix(16, 0) = "Welcome Basket"
    grdAmen.TextMatrix(17, 0) = "Wireless Hotspot"
    grdRooms.Rows = 1
    grdRooms.ColWidth(0) = grdRooms.Width * 0.12
    grdRooms.ColWidth(1) = grdRooms.Width * 0.4
    grdRooms.ColWidth(2) = grdRooms.Width * 0.12
    grdRooms.ColWidth(3) = grdRooms.Width * 0.12
    grdRooms.ColWidth(4) = grdRooms.Width * 0.22
    grdRooms.TextMatrix(0, 0) = "Room Number"
    grdRooms.TextMatrix(0, 1) = "Description"
    grdRooms.TextMatrix(0, 2) = "Number of Beds"
    grdRooms.TextMatrix(0, 3) = "Number of Baths"
    grdRooms.TextMatrix(0, 4) = "Rate"
    Set Picture1.Picture = ImageList1.ListImages(11).Picture
    Set Picture2.Picture = ImageList1.ListImages(4).Picture
    LoadRooms
    If grdRooms.Rows = 1 Then
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = False
        frmMain.Toolbar1.Buttons(5).Enabled = False
    Else
        grdRooms.Row = 1
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(5).Enabled = True
    End If
    grdRooms.ColAlignment(0) = flexAlignLeftCenter
    grdRooms.ColAlignment(1) = flexAlignLeftCenter
    grdRooms.ColAlignment(2) = flexAlignLeftCenter
    grdRooms.ColAlignment(3) = flexAlignLeftCenter
    grdRooms.ColAlignment(4) = flexAlignLeftCenter
    Exit Sub
trap:
    If err.Number = 76 Then
        MkDir App.Path & "\Rooms"
    End If
    Resume Next
End Sub
Private Sub LoadRooms()
    picLocation.Visible = False
    cmdLocation.Value = Up
    grdRooms.Rows = 1
    ActiveReadServer "Select * from Rooms order by Room_No"
    grdRooms.Rows = rs.RecordCount + 1
    i = 0
    While Not rs.EOF
        i = i + 1
        grdRooms.TextMatrix(i, 0) = rs.Fields("Room_No")
        grdRooms.TextMatrix(i, 1) = rs.Fields("Description")
        grdRooms.TextMatrix(i, 2) = rs.Fields("Bed_No")
        grdRooms.TextMatrix(i, 3) = rs.Fields("Bath_No")
        If Val(rs.Fields("Room_Rate") & "") = 0 Then
            grdRooms.TextMatrix(i, 4) = "<Select a Rate>"
        Else
            ActiveReadServer1 "Select * from Room_Rates where Rate_Type = " & Val(rs.Fields("Room_Rate") & "")
            If rs1.RecordCount > 0 Then
                grdRooms.TextMatrix(i, 4) = rs1.Fields("Rate_Type") & " - " & rs1.Fields("Description") & " > " & Format(rs1.Fields("Room_Rate"), "0.00")
            Else
                grdRooms.TextMatrix(i, 4) = "<Select a Rate>"
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    If grdRooms.Rows > 1 Then grdRooms.Row = 1
End Sub

Private Sub Form_LostFocus()
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
End Sub
Private Sub grdAmen_EnterCell()
    Select Case grdAmen.Col
        Case 0
            grdAmen.Editable = flexEDNone
        Case 1
            grdAmen.Editable = flexEDKbdMouse
    End Select
End Sub

Private Sub grdAmen_GotFocus()
    picLocation.Visible = False
End Sub

Private Sub grdLocations_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    grdLocations.Tag = ""
    For i = 1 To grdLocations.Rows - 1
        If i <> Row Then
            grdLocations.TextMatrix(i, 2) = ""
        End If
        If grdLocations.TextMatrix(i, 2) <> "" Then
            grdLocations.Tag = Val(grdLocations.TextMatrix(i, 0))
        End If
    Next i
    If cmbRoom_No.Text <> "" And txtDescription.Text <> "" Then
         frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub grdLocations_EnterCell()
    If grdLocations.Col <> 2 Then
        grdLocations.Col = 2
    End If
    grdLocations.Editable = flexEDKbdMouse
End Sub

Private Sub grdRooms_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
DoEvents
Dim Amen As Variant
On Error Resume Next
    For i = 1 To grdAmen.Rows - 1
        grdAmen.TextMatrix(i, 1) = "0"
    Next i
    If grdRooms.TextMatrix(NewRow, 0) <> "" And NewRow <> 0 Then
        ActiveReadServer "Select * from Rooms where Room_No = " & grdRooms.TextMatrix(NewRow, 0)
        If rs.RecordCount > 0 Then
            cmbRoom_No.Text = rs.Fields("Room_No")
            txtDescription.Text = rs.Fields("Description")
            txtKing.Text = rs.Fields("King_Size")
            txtQueen.Text = rs.Fields("Queen_Size")
            txtDouble.Text = rs.Fields("Double_Bed")
            txtSingle.Text = rs.Fields("Single_Bed")
            txtBaby.Text = rs.Fields("Baby_Bed")
            txtStraightbath.Text = rs.Fields("Straight")
            txtCorner.Text = rs.Fields("Corner")
            txtRoll.Text = rs.Fields("Roll_Top")
            txtShowerb.Text = rs.Fields("Shower_Bath")
            txtSpa.Text = rs.Fields("Spa")
            txtShower1.Text = rs.Fields("Shower")
            cmbRates.Text = grdRooms.TextMatrix(NewRow, 4)
            Amen = Split(rs.Fields("Ameneties"), "|")
            For i = 0 To UBound(Amen)
                grdAmen.TextMatrix(Amen(i), 1) = "1"
            Next i
            If Val(rs.Fields("Location_no") & "") <> 0 Then
                For i = 1 To grdLocations.Rows - 1
                    If grdLocations.ValueMatrix(i, 0) = Val(rs.Fields("Location_no")) & "" Then
                        grdLocations.TextMatrix(i, 2) = "1"
                    Else
                        grdLocations.TextMatrix(i, 2) = ""
                    End If
                Next i
            End If
            If Trim(rs.Fields("Room_File") & "") <> "" Then
                Label4.Visible = False
                Set picRooms.Picture = LoadPicture(rs.Fields("Room_File"))
            Else
                Label4.Visible = True
                Set picRooms.Picture = LoadPicture("")
                
            End If
            frmRooms.Refresh
        End If
    End If
On Error GoTo 0
End Sub

Private Sub grdRooms_AfterSort(ByVal Col As Long, Order As Integer)
    Dim Amen As Variant
On Error Resume Next
    For i = 1 To grdAmen.Rows - 1
        grdAmen.TextMatrix(i, 1) = "0"
    Next i
    If grdRooms.TextMatrix(NewRow, 0) <> "" And NewRow <> 0 Then
        ActiveReadServer "Select * from Rooms where Room_No = " & grdRooms.TextMatrix(NewRow, 0)
        If rs.RecordCount > 0 Then
            cmbRoom_No.Text = rs.Fields("Room_No")
            txtDescription.Text = rs.Fields("Description")
            txtKing.Text = rs.Fields("King_Size")
            txtQueen.Text = rs.Fields("Queen_Size")
            txtDouble.Text = rs.Fields("Double_Bed")
            txtSingle.Text = rs.Fields("Single_Bed")
            txtBaby.Text = rs.Fields("Baby_Bed")
            txtStraightbath.Text = rs.Fields("Straight")
            txtCorner.Text = rs.Fields("Corner")
            txtRoll.Text = rs.Fields("Roll_Top")
            txtShowerb.Text = rs.Fields("Shower_Bath")
            txtSpa.Text = rs.Fields("Spa")
            txtShower1.Text = rs.Fields("Shower")
            If Val(rs.Fields("Room_Rate") & "") = 0 Then
                cmbRates.Text = "<Select a Rate>"
            Else
                ActiveReadServer1 "Select * from Room_Rates where Rate_Type = " & Val(rs.Fields("Room_Rate") & "")
                If rs1.RecordCount > 0 Then
                    cmbRates.Text = rs1.Fields("Rate_Type") & " - " & rs1.Fields("Description") & " > " & Format(rs1.Fields("Room_Rate"), "0.00")
                Else
                    cmbRates.Text = "<Select a Rate>"
                End If
            End If
            Amen = Split(rs.Fields("Ameneties"), "|")
            For i = 0 To UBound(Amen)
                grdAmen.TextMatrix(Amen(i), 1) = "1"
            Next i
            If Val(rs.Fields("Location_no") & "") <> 0 Then
                For i = 1 To grdLocations.Rows - 1
                    If grdLocations.ValueMatrix(i, 0) = Val(rs.Fields("Location_no")) & "" Then
                        grdLocations.TextMatrix(i, 2) = "1"
                    Else
                        grdLocations.TextMatrix(i, 2) = ""
                    End If
                Next i
            End If
            If rs.Fields("Room_File") <> "" Then
                Label4.Visible = False
                Set picRooms.Picture = LoadPicture(rs.Fields("Room_File"))
            Else
                Label4.Visible = True
                Set picRooms.Picture = LoadPicture("")
            End If
        End If
    End If
On Error GoTo 0
End Sub
Private Sub Load_Logo(Action)
    cmLogo.Action = 1
    cmLogo.InitDir = App.Path & "\Images"
    Select Case Action
        Case 0
            Label4.Visible = False
            Set picRooms.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set picRooms.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Room_File = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 1
            Label5.Visible = False
            Set Picture2.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture2.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic1 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 2
            Label5.Visible = False
            Set Picture2.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture2.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic2 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 3
            Label5.Visible = False
            Set Picture2.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture2.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic3 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 4
            Label5.Visible = False
            Set Picture2.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture2.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic4 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 5
            Label5.Visible = False
            Set Picture2.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture2.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic5 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 6
            Label5.Visible = False
            Set Picture1.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture1.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic6 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 7
            Label5.Visible = False
            Set Picture1.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture1.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic7 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 8
            Label5.Visible = False
            Set Picture1.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture1.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic8 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 9
            Label5.Visible = False
            Set Picture1.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture1.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic9 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 10
            Label5.Visible = False
            Set Picture1.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture1.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic10 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
        Case 11
            Label5.Visible = False
            Set Picture1.Picture = LoadPicture("")
            DoEvents
            Logo_File = cmLogo.FileName
            Set Picture1.Picture = LoadPicture(cmLogo.FileName)
            ActiveUpdateServer "Update Rooms set Pic11 = '" & cmLogo.FileName & "' where Room_No = " & cmbRoom_No.Text
    End Select
    frmRooms.Refresh
End Sub

Private Sub grdRooms_GotFocus()
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub Image7_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub Label2_Click(Index As Integer)
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub Label4_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    Load_Logo 0
End Sub

Private Sub Label5_Click()
    Load_Logo 1
End Sub

Private Sub Label6_Click()
    Load_Logo 2
End Sub

Private Sub lblRoomDet_Click(Index As Integer)
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub picRoom_Click()
    Load_Logo 0
End Sub

Private Sub picRooms_Click()
    picLocation.Visible = False
    cmdLocation.Value = Up
    Load_Logo 0
End Sub
Private Sub Picture1_Click()
    Load_Logo Picture1.Tag
End Sub
Private Sub Picture2_Click()
    Load_Logo Picture2.Tag
End Sub
Private Sub txtBaby_GotFocus()
    txtBaby.SelStart = 0
    txtBaby.SelLength = Len(txtBaby.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub
Private Sub txtBaby_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtSingle.SetFocus
        Case 40: KeyCode = 0: txtStraightbath.SetFocus
    End Select
End Sub
Private Sub txtBaby_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtBaby_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtcorner_GotFocus()
    txtCorner.SelStart = 0
    txtCorner.SelLength = Len(txtCorner.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub
Private Sub txtcorner_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtStraightbath.SetFocus
        Case 40: KeyCode = 0: txtRoll.SetFocus
    End Select
End Sub

Private Sub txtcorner_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcorner_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtDescription_Change()
    If cmbRoom_No.Text <> "" And txtDescription.Text <> "" Then
         frmMain.Toolbar1.Buttons(4).Enabled = True
    Else
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Private Sub txtDescription_GotFocus()
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: cmbRoom_No.SetFocus
        Case 40: KeyCode = 0: txtKing.SetFocus
    End Select
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 32
        Case 65 To 90, 97 To 122
    End Select
End Sub

Private Sub txtDescription_LostFocus()
    On Error Resume Next
    txtDescription.Text = UCase(Left(txtDescription.Text, 1)) & Mid(txtDescription.Text, 2)
    On Error GoTo 0
End Sub

Private Sub txtDouble_GotFocus()
    txtDouble.SelStart = 0
    txtDouble.SelLength = Len(txtDouble.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtDouble_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtQueen.SetFocus
        Case 40: KeyCode = 0: txtSingle.SetFocus
    End Select
End Sub

Private Sub txtDouble_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtDouble_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtKing_GotFocus()
    txtKing.SelStart = 0
    txtKing.SelLength = Len(txtKing.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtKing_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtDescription.SetFocus
        Case 40: KeyCode = 0: txtQueen.SetFocus
    End Select
End Sub

Private Sub txtking_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtKing_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtQueen_GotFocus()
    txtQueen.SelStart = 0
    txtQueen.SelLength = Len(txtQueen.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub
Private Sub txtQueen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtKing.SetFocus
        Case 40: KeyCode = 0: txtDouble.SetFocus
    End Select
End Sub

Private Sub txtqueen_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtQueen_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtroll_GotFocus()
    txtRoll.SelStart = 0
    txtRoll.SelLength = Len(txtRoll.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtroll_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtCorner.SetFocus
        Case 40: KeyCode = 0: txtShowerb.SetFocus
    End Select
End Sub

Private Sub txtroll_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtroll_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtshower1_GotFocus()
    txtShower1.SelStart = 0
    txtShower1.SelLength = Len(txtShower1.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtshower1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtSpa.SetFocus
        Case 40: KeyCode = 0: cmbRates.SetFocus
    End Select
End Sub

Private Sub txtshower1_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtshower1_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtshowerb_GotFocus()
    txtShowerb.SelStart = 0
    txtShowerb.SelLength = Len(txtShowerb.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtshowerb_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtRoll.SetFocus
        Case 40: KeyCode = 0: txtSpa.SetFocus
    End Select
End Sub

Private Sub txtshowerb_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtshowerb_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtSingle_GotFocus()
    txtSingle.SelStart = 0
    txtSingle.SelLength = Len(txtSingle.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtSingle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtDouble.SetFocus
        Case 40: KeyCode = 0: txtBaby.SetFocus
    End Select
End Sub

Private Sub txtSingle_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSingle_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtSpa_GotFocus()
    txtSpa.SelStart = 0
    txtSpa.SelLength = Len(txtSpa.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtSpa_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtShowerb.SetFocus
        Case 40: KeyCode = 0: txtShower1.SetFocus
    End Select
End Sub

Private Sub txtSpa_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSpa_LostFocus()
    picRoom.Visible = True
End Sub

Private Sub txtStraightbath_GotFocus()
    txtStraightbath.SelStart = 0
    txtStraightbath.SelLength = Len(txtStraightbath.Text)
    picRoom.Visible = False
    picLocation.Visible = False
    cmdLocation.Value = Up
End Sub

Private Sub txtStraightbath_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: KeyCode = 0: txtBaby.SetFocus
        Case 40: KeyCode = 0: txtCorner.SetFocus
    End Select
End Sub

Private Sub txtStraightbath_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case 8
        Case 39
            KeyAscii = 0
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtStraightbath_LostFocus()
    picRoom.Visible = True
End Sub
