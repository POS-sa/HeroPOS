VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmRequest 
   ClientHeight    =   11400
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H000000C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmRequest.frx":0000
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid grdPlu 
      Height          =   5910
      Left            =   150
      TabIndex        =   59
      Top             =   720
      Visible         =   0   'False
      Width           =   105
      _cx             =   185
      _cy             =   10425
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   15790305
      BackColorAlternate=   16381166
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
      Rows            =   15
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   1500
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
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
   Begin VB.Timer ScrolTimer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   0
      Top             =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid grdLoc 
      Height          =   45
      Left            =   4560
      TabIndex        =   58
      Top             =   1260
      Visible         =   0   'False
      Width           =   6285
      _cx             =   11086
      _cy             =   79
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   11852525
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   3119822
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   11852525
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
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   650
      RowHeightMax    =   0
      ColWidthMin     =   3500
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
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
      WallPaper       =   "frmRequest.frx":10207
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picSlip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10785
      Left            =   4440
      ScaleHeight     =   10755
      ScaleWidth      =   6450
      TabIndex        =   62
      Top             =   360
      Visible         =   0   'False
      Width           =   6480
      Begin VSFlex8Ctl.VSFlexGrid grdRem 
         Height          =   8880
         Left            =   90
         TabIndex        =   66
         Top             =   690
         Width           =   5505
         _cx             =   9710
         _cy             =   15663
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   12377839
         ForeColor       =   -2147483640
         BackColorFixed  =   7848417
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4827091
         ForeColorSel    =   0
         BackColorBkg    =   7453147
         BackColorAlternate=   12377839
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   630
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
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
         WallPaper       =   "frmRequest.frx":13F38
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin btButtonEx.ButtonEx cmdClose 
         Height          =   990
         Left            =   4230
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   9660
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1746
         Appearance      =   3
         BackColor       =   6208225
         BorderColor     =   8421504
         Caption         =   "Close"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   660
         Index           =   1
         Left            =   5600
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   8910
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1164
         Appearance      =   3
         BackColor       =   6208225
         BorderColor     =   1876185
         Caption         =   "6"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   660
         Index           =   0
         Left            =   5600
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   690
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1164
         Appearance      =   3
         BackColor       =   6208225
         BorderColor     =   1876185
         Caption         =   "5"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdRemove 
         Height          =   990
         Left            =   90
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   9660
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1746
         Appearance      =   3
         BackColor       =   2720171
         BorderColor     =   8421504
         Caption         =   "Remove"
         CaptionOffsetX  =   1
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin VB.Shape picGrid 
         BackColor       =   &H00B4DAED&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0071B9DB&
         FillStyle       =   2  'Horizontal Line
         Height          =   8265
         Left            =   5610
         Top             =   705
         Width           =   735
      End
      Begin MSForms.Label lblHead 
         Height          =   435
         Left            =   240
         TabIndex        =   67
         Top             =   165
         Width           =   5985
         ForeColor       =   7555868
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Size            =   "10557;767"
         FontName        =   "Arial Narrow"
         FontHeight      =   315
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image Image7 
         Height          =   525
         Left            =   90
         Top             =   110
         Width           =   6265
         BackColor       =   11852525
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "11051;926"
      End
      Begin MSForms.Image Image6 
         Height          =   8865
         Left            =   90
         Top             =   690
         Width           =   6255
         BorderStyle     =   0
         SpecialEffect   =   2
         Size            =   "11033;15637"
      End
      Begin MSForms.Image Image5 
         Height          =   10755
         Left            =   -240
         Top             =   0
         Width           =   6705
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "11827;18971"
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdDept 
      Height          =   8070
      Left            =   60
      TabIndex        =   54
      Top             =   930
      Visible         =   0   'False
      Width           =   195
      _cx             =   344
      _cy             =   14235
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   15790305
      BackColorAlternate=   16381166
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
      Rows            =   15
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   1500
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
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
   Begin VB.Timer scrolTimer 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   420
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9420
      Top             =   270
   End
   Begin BTNENHLib4.BtnEnh cmdLogoff 
      Height          =   2005
      Left            =   390
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   1905
      _Version        =   524298
      _ExtentX        =   3360
      _ExtentY        =   3537
      _StockProps     =   66
      Caption         =   "Clear Request"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":15DDB
      textLT          =   "frmRequest.frx":15E55
      textCT          =   "frmRequest.frx":15E6D
      textRT          =   "frmRequest.frx":15E85
      textLM          =   "frmRequest.frx":15E9D
      textRM          =   "frmRequest.frx":15EB5
      textLB          =   "frmRequest.frx":15ECD
      textCB          =   "frmRequest.frx":15EE5
      textRB          =   "frmRequest.frx":15EFD
      colorBack       =   "frmRequest.frx":15F15
      colorIntern     =   "frmRequest.frx":15F3F
      colorMO         =   "frmRequest.frx":15F69
      colorFocus      =   "frmRequest.frx":15F93
      colorDisabled   =   "frmRequest.frx":15FBD
      colorPressed    =   "frmRequest.frx":15FE7
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   960
      Left            =   8040
      TabIndex        =   1
      Top             =   300
      Visible         =   0   'False
      Width           =   4935
      _Version        =   524298
      _ExtentX        =   8705
      _ExtentY        =   1693
      _StockProps     =   66
      Caption         =   "Invalid Key Pressed"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   15
      BackColorContainer=   12632256
      SpecialEffect   =   3
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":16011
      textLT          =   "frmRequest.frx":16097
      textCT          =   "frmRequest.frx":160AF
      textRT          =   "frmRequest.frx":160C7
      textLM          =   "frmRequest.frx":160DF
      textRM          =   "frmRequest.frx":160F7
      textLB          =   "frmRequest.frx":1610F
      textCB          =   "frmRequest.frx":16127
      textRB          =   "frmRequest.frx":1613F
      colorBack       =   "frmRequest.frx":16157
      colorIntern     =   "frmRequest.frx":16181
      colorMO         =   "frmRequest.frx":161AB
      colorFocus      =   "frmRequest.frx":161D5
      colorDisabled   =   "frmRequest.frx":161FF
      colorPressed    =   "frmRequest.frx":16229
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   0
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2745
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":16253
      textLT          =   "frmRequest.frx":1626B
      textCT          =   "frmRequest.frx":16283
      textRT          =   "frmRequest.frx":1629B
      textLM          =   "frmRequest.frx":162B3
      textRM          =   "frmRequest.frx":162CB
      textLB          =   "frmRequest.frx":162E3
      textCB          =   "frmRequest.frx":162FB
      textRB          =   "frmRequest.frx":16313
      colorBack       =   "frmRequest.frx":1632B
      colorIntern     =   "frmRequest.frx":16355
      colorMO         =   "frmRequest.frx":1637F
      colorFocus      =   "frmRequest.frx":163A9
      colorDisabled   =   "frmRequest.frx":163D3
      colorPressed    =   "frmRequest.frx":163FD
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   3
      Left            =   390
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4050
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":16427
      textLT          =   "frmRequest.frx":1643F
      textCT          =   "frmRequest.frx":16457
      textRT          =   "frmRequest.frx":1646F
      textLM          =   "frmRequest.frx":16487
      textRM          =   "frmRequest.frx":1649F
      textLB          =   "frmRequest.frx":164B7
      textCB          =   "frmRequest.frx":164CF
      textRB          =   "frmRequest.frx":164E7
      colorBack       =   "frmRequest.frx":164FF
      colorIntern     =   "frmRequest.frx":16529
      colorMO         =   "frmRequest.frx":16553
      colorFocus      =   "frmRequest.frx":1657D
      colorDisabled   =   "frmRequest.frx":165A7
      colorPressed    =   "frmRequest.frx":165D1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   5
      Left            =   5100
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4050
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      PicturePosition =   9
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":165FB
      textLT          =   "frmRequest.frx":16613
      textCT          =   "frmRequest.frx":1662B
      textRT          =   "frmRequest.frx":16643
      textLM          =   "frmRequest.frx":1665B
      textRM          =   "frmRequest.frx":16673
      textLB          =   "frmRequest.frx":1668B
      textCB          =   "frmRequest.frx":166A3
      textRB          =   "frmRequest.frx":166BB
      colorBack       =   "frmRequest.frx":166D3
      colorIntern     =   "frmRequest.frx":166FD
      colorMO         =   "frmRequest.frx":16727
      colorFocus      =   "frmRequest.frx":16751
      colorDisabled   =   "frmRequest.frx":1677B
      colorPressed    =   "frmRequest.frx":167A5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   6
      Left            =   390
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":167CF
      textLT          =   "frmRequest.frx":167E7
      textCT          =   "frmRequest.frx":167FF
      textRT          =   "frmRequest.frx":16817
      textLM          =   "frmRequest.frx":1682F
      textRM          =   "frmRequest.frx":16847
      textLB          =   "frmRequest.frx":1685F
      textCB          =   "frmRequest.frx":16877
      textRB          =   "frmRequest.frx":1688F
      colorBack       =   "frmRequest.frx":168A7
      colorIntern     =   "frmRequest.frx":168D1
      colorMO         =   "frmRequest.frx":168FB
      colorFocus      =   "frmRequest.frx":16925
      colorDisabled   =   "frmRequest.frx":1694F
      colorPressed    =   "frmRequest.frx":16979
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   8
      Left            =   5100
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":169A3
      textLT          =   "frmRequest.frx":169BB
      textCT          =   "frmRequest.frx":169D3
      textRT          =   "frmRequest.frx":169EB
      textLM          =   "frmRequest.frx":16A03
      textRM          =   "frmRequest.frx":16A1B
      textLB          =   "frmRequest.frx":16A33
      textCB          =   "frmRequest.frx":16A4B
      textRB          =   "frmRequest.frx":16A63
      colorBack       =   "frmRequest.frx":16A7B
      colorIntern     =   "frmRequest.frx":16AA5
      colorMO         =   "frmRequest.frx":16ACF
      colorFocus      =   "frmRequest.frx":16AF9
      colorDisabled   =   "frmRequest.frx":16B23
      colorPressed    =   "frmRequest.frx":16B4D
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   9
      Left            =   390
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6660
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":16B77
      textLT          =   "frmRequest.frx":16B8F
      textCT          =   "frmRequest.frx":16BA7
      textRT          =   "frmRequest.frx":16BBF
      textLM          =   "frmRequest.frx":16BD7
      textRM          =   "frmRequest.frx":16BEF
      textLB          =   "frmRequest.frx":16C07
      textCB          =   "frmRequest.frx":16C1F
      textRB          =   "frmRequest.frx":16C37
      colorBack       =   "frmRequest.frx":16C4F
      colorIntern     =   "frmRequest.frx":16C79
      colorMO         =   "frmRequest.frx":16CA3
      colorFocus      =   "frmRequest.frx":16CCD
      colorDisabled   =   "frmRequest.frx":16CF7
      colorPressed    =   "frmRequest.frx":16D21
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   11
      Left            =   5100
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6660
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      PicturePosition =   9
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":16D4B
      textLT          =   "frmRequest.frx":16D63
      textCT          =   "frmRequest.frx":16D7B
      textRT          =   "frmRequest.frx":16D93
      textLM          =   "frmRequest.frx":16DAB
      textRM          =   "frmRequest.frx":16DC3
      textLB          =   "frmRequest.frx":16DDB
      textCB          =   "frmRequest.frx":16DF3
      textRB          =   "frmRequest.frx":16E0B
      colorBack       =   "frmRequest.frx":16E23
      colorIntern     =   "frmRequest.frx":16E4D
      colorMO         =   "frmRequest.frx":16E77
      colorFocus      =   "frmRequest.frx":16EA1
      colorDisabled   =   "frmRequest.frx":16ECB
      colorPressed    =   "frmRequest.frx":16EF5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   12
      Left            =   390
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7950
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":16F1F
      textLT          =   "frmRequest.frx":16F37
      textCT          =   "frmRequest.frx":16F4F
      textRT          =   "frmRequest.frx":16F67
      textLM          =   "frmRequest.frx":16F7F
      textRM          =   "frmRequest.frx":16F97
      textLB          =   "frmRequest.frx":16FAF
      textCB          =   "frmRequest.frx":16FC7
      textRB          =   "frmRequest.frx":16FDF
      colorBack       =   "frmRequest.frx":16FF7
      colorIntern     =   "frmRequest.frx":17021
      colorMO         =   "frmRequest.frx":1704B
      colorFocus      =   "frmRequest.frx":17075
      colorDisabled   =   "frmRequest.frx":1709F
      colorPressed    =   "frmRequest.frx":170C9
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   14
      Left            =   5100
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7965
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":170F3
      textLT          =   "frmRequest.frx":1710B
      textCT          =   "frmRequest.frx":17123
      textRT          =   "frmRequest.frx":1713B
      textLM          =   "frmRequest.frx":17153
      textRM          =   "frmRequest.frx":1716B
      textLB          =   "frmRequest.frx":17183
      textCB          =   "frmRequest.frx":1719B
      textRB          =   "frmRequest.frx":171B3
      colorBack       =   "frmRequest.frx":171CB
      colorIntern     =   "frmRequest.frx":171F5
      colorMO         =   "frmRequest.frx":1721F
      colorFocus      =   "frmRequest.frx":17249
      colorDisabled   =   "frmRequest.frx":17273
      colorPressed    =   "frmRequest.frx":1729D
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   15
      Left            =   390
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9270
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":172C7
      textLT          =   "frmRequest.frx":172DF
      textCT          =   "frmRequest.frx":172F7
      textRT          =   "frmRequest.frx":1730F
      textLM          =   "frmRequest.frx":17327
      textRM          =   "frmRequest.frx":1733F
      textLB          =   "frmRequest.frx":17357
      textCB          =   "frmRequest.frx":1736F
      textRB          =   "frmRequest.frx":17387
      colorBack       =   "frmRequest.frx":1739F
      colorIntern     =   "frmRequest.frx":173C9
      colorMO         =   "frmRequest.frx":173F3
      colorFocus      =   "frmRequest.frx":1741D
      colorDisabled   =   "frmRequest.frx":17447
      colorPressed    =   "frmRequest.frx":17471
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   17
      Left            =   5100
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9270
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1749B
      textLT          =   "frmRequest.frx":174B3
      textCT          =   "frmRequest.frx":174CB
      textRT          =   "frmRequest.frx":174E3
      textLM          =   "frmRequest.frx":174FB
      textRM          =   "frmRequest.frx":17513
      textLB          =   "frmRequest.frx":1752B
      textCB          =   "frmRequest.frx":17543
      textRB          =   "frmRequest.frx":1755B
      colorBack       =   "frmRequest.frx":17573
      colorIntern     =   "frmRequest.frx":1759D
      colorMO         =   "frmRequest.frx":175C7
      colorFocus      =   "frmRequest.frx":175F1
      colorDisabled   =   "frmRequest.frx":1761B
      colorPressed    =   "frmRequest.frx":17645
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   1
      Left            =   2745
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2745
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1766F
      textLT          =   "frmRequest.frx":17687
      textCT          =   "frmRequest.frx":1769F
      textRT          =   "frmRequest.frx":176B7
      textLM          =   "frmRequest.frx":176CF
      textRM          =   "frmRequest.frx":176E7
      textLB          =   "frmRequest.frx":176FF
      textCB          =   "frmRequest.frx":17717
      textRB          =   "frmRequest.frx":1772F
      colorBack       =   "frmRequest.frx":17747
      colorIntern     =   "frmRequest.frx":17771
      colorMO         =   "frmRequest.frx":1779B
      colorFocus      =   "frmRequest.frx":177C5
      colorDisabled   =   "frmRequest.frx":177EF
      colorPressed    =   "frmRequest.frx":17819
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   4
      Left            =   2745
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4050
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      PicturePosition =   9
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":17843
      textLT          =   "frmRequest.frx":1785B
      textCT          =   "frmRequest.frx":17873
      textRT          =   "frmRequest.frx":1788B
      textLM          =   "frmRequest.frx":178A3
      textRM          =   "frmRequest.frx":178BB
      textLB          =   "frmRequest.frx":178D3
      textCB          =   "frmRequest.frx":178EB
      textRB          =   "frmRequest.frx":17903
      colorBack       =   "frmRequest.frx":1791B
      colorIntern     =   "frmRequest.frx":17945
      colorMO         =   "frmRequest.frx":1796F
      colorFocus      =   "frmRequest.frx":17999
      colorDisabled   =   "frmRequest.frx":179C3
      colorPressed    =   "frmRequest.frx":179ED
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   7
      Left            =   2745
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":17A17
      textLT          =   "frmRequest.frx":17A2F
      textCT          =   "frmRequest.frx":17A47
      textRT          =   "frmRequest.frx":17A5F
      textLM          =   "frmRequest.frx":17A77
      textRM          =   "frmRequest.frx":17A8F
      textLB          =   "frmRequest.frx":17AA7
      textCB          =   "frmRequest.frx":17ABF
      textRB          =   "frmRequest.frx":17AD7
      colorBack       =   "frmRequest.frx":17AEF
      colorIntern     =   "frmRequest.frx":17B19
      colorMO         =   "frmRequest.frx":17B43
      colorFocus      =   "frmRequest.frx":17B6D
      colorDisabled   =   "frmRequest.frx":17B97
      colorPressed    =   "frmRequest.frx":17BC1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   10
      Left            =   2745
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6660
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      PicturePosition =   9
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":17BEB
      textLT          =   "frmRequest.frx":17C03
      textCT          =   "frmRequest.frx":17C1B
      textRT          =   "frmRequest.frx":17C33
      textLM          =   "frmRequest.frx":17C4B
      textRM          =   "frmRequest.frx":17C63
      textLB          =   "frmRequest.frx":17C7B
      textCB          =   "frmRequest.frx":17C93
      textRB          =   "frmRequest.frx":17CAB
      colorBack       =   "frmRequest.frx":17CC3
      colorIntern     =   "frmRequest.frx":17CED
      colorMO         =   "frmRequest.frx":17D17
      colorFocus      =   "frmRequest.frx":17D41
      colorDisabled   =   "frmRequest.frx":17D6B
      colorPressed    =   "frmRequest.frx":17D95
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   13
      Left            =   2745
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7965
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":17DBF
      textLT          =   "frmRequest.frx":17DD7
      textCT          =   "frmRequest.frx":17DEF
      textRT          =   "frmRequest.frx":17E07
      textLM          =   "frmRequest.frx":17E1F
      textRM          =   "frmRequest.frx":17E37
      textLB          =   "frmRequest.frx":17E4F
      textCB          =   "frmRequest.frx":17E67
      textRB          =   "frmRequest.frx":17E7F
      colorBack       =   "frmRequest.frx":17E97
      colorIntern     =   "frmRequest.frx":17EC1
      colorMO         =   "frmRequest.frx":17EEB
      colorFocus      =   "frmRequest.frx":17F15
      colorDisabled   =   "frmRequest.frx":17F3F
      colorPressed    =   "frmRequest.frx":17F69
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   16
      Left            =   2745
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9270
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":17F93
      textLT          =   "frmRequest.frx":17FAB
      textCT          =   "frmRequest.frx":17FC3
      textRT          =   "frmRequest.frx":17FDB
      textLM          =   "frmRequest.frx":17FF3
      textRM          =   "frmRequest.frx":1800B
      textLB          =   "frmRequest.frx":18023
      textCB          =   "frmRequest.frx":1803B
      textRB          =   "frmRequest.frx":18053
      colorBack       =   "frmRequest.frx":1806B
      colorIntern     =   "frmRequest.frx":18095
      colorMO         =   "frmRequest.frx":180BF
      colorFocus      =   "frmRequest.frx":180E9
      colorDisabled   =   "frmRequest.frx":18113
      colorPressed    =   "frmRequest.frx":1813D
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   4545
      Left            =   10830
      TabIndex        =   19
      Top             =   1485
      Width           =   4215
      _cx             =   7435
      _cy             =   8017
      Appearance      =   0
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
      BackColor       =   12377839
      ForeColor       =   -2147483640
      BackColorFixed  =   7848417
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4827091
      ForeColorSel    =   0
      BackColorBkg    =   7453147
      BackColorAlternate=   12377839
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
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
      WallPaper       =   "frmRequest.frx":18167
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   2025
      Index           =   3
      Left            =   2250
      TabIndex        =   20
      Top             =   525
      Width           =   2235
      _Version        =   524298
      _ExtentX        =   3942
      _ExtentY        =   3572
      _StockProps     =   66
      Caption         =   "Place Request"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
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
      Shape           =   4
      CornerFactor    =   80
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      CaptionWordWrapPerc=   80
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1A00A
      textLT          =   "frmRequest.frx":1A084
      textCT          =   "frmRequest.frx":1A09C
      textRT          =   "frmRequest.frx":1A0B4
      textLM          =   "frmRequest.frx":1A0CC
      textRM          =   "frmRequest.frx":1A0E4
      textLB          =   "frmRequest.frx":1A0FC
      textCB          =   "frmRequest.frx":1A114
      textRB          =   "frmRequest.frx":1A12C
      colorBack       =   "frmRequest.frx":1A144
      colorIntern     =   "frmRequest.frx":1A16E
      colorMO         =   "frmRequest.frx":1A198
      colorFocus      =   "frmRequest.frx":1A1C2
      colorDisabled   =   "frmRequest.frx":1A1EC
      colorPressed    =   "frmRequest.frx":1A216
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1305
      Index           =   2
      Left            =   5100
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2745
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   2302
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   10
      Surface         =   1
      BackColorContainer=   3119822
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1A240
      textLT          =   "frmRequest.frx":1A258
      textCT          =   "frmRequest.frx":1A270
      textRT          =   "frmRequest.frx":1A288
      textLM          =   "frmRequest.frx":1A2A0
      textRM          =   "frmRequest.frx":1A2B8
      textLB          =   "frmRequest.frx":1A2D0
      textCB          =   "frmRequest.frx":1A2E8
      textRB          =   "frmRequest.frx":1A300
      colorBack       =   "frmRequest.frx":1A318
      colorIntern     =   "frmRequest.frx":1A342
      colorMO         =   "frmRequest.frx":1A36C
      colorFocus      =   "frmRequest.frx":1A396
      colorDisabled   =   "frmRequest.frx":1A3C0
      colorPressed    =   "frmRequest.frx":1A3EA
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1125
      Index           =   6
      Left            =   4500
      TabIndex        =   22
      Top             =   1425
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
      _ExtentY        =   1984
      _StockProps     =   66
      Caption         =   "Remove Item"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   20
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1A414
      textLT          =   "frmRequest.frx":1A48A
      textCT          =   "frmRequest.frx":1A4A2
      textRT          =   "frmRequest.frx":1A4BA
      textLM          =   "frmRequest.frx":1A4D2
      textRM          =   "frmRequest.frx":1A4EA
      textLB          =   "frmRequest.frx":1A502
      textCB          =   "frmRequest.frx":1A51A
      textRB          =   "frmRequest.frx":1A532
      colorBack       =   "frmRequest.frx":1A54A
      colorIntern     =   "frmRequest.frx":1A574
      colorMO         =   "frmRequest.frx":1A59E
      colorFocus      =   "frmRequest.frx":1A5C8
      colorDisabled   =   "frmRequest.frx":1A5F2
      colorPressed    =   "frmRequest.frx":1A61C
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   3
      Left            =   9090
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1A646
      textLT          =   "frmRequest.frx":1A6A8
      textCT          =   "frmRequest.frx":1A6C0
      textRT          =   "frmRequest.frx":1A6D8
      textLM          =   "frmRequest.frx":1A6F0
      textRM          =   "frmRequest.frx":1A708
      textLB          =   "frmRequest.frx":1A720
      textCB          =   "frmRequest.frx":1A738
      textRB          =   "frmRequest.frx":1A750
      colorBack       =   "frmRequest.frx":1A768
      colorIntern     =   "frmRequest.frx":1A792
      colorMO         =   "frmRequest.frx":1A7BC
      colorFocus      =   "frmRequest.frx":1A7E6
      colorDisabled   =   "frmRequest.frx":1A810
      colorPressed    =   "frmRequest.frx":1A83A
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   5
      Left            =   9090
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3675
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1A864
      textLT          =   "frmRequest.frx":1A8C6
      textCT          =   "frmRequest.frx":1A8DE
      textRT          =   "frmRequest.frx":1A8F6
      textLM          =   "frmRequest.frx":1A90E
      textRM          =   "frmRequest.frx":1A926
      textLB          =   "frmRequest.frx":1A93E
      textCB          =   "frmRequest.frx":1A956
      textRB          =   "frmRequest.frx":1A96E
      colorBack       =   "frmRequest.frx":1A986
      colorIntern     =   "frmRequest.frx":1A9B0
      colorMO         =   "frmRequest.frx":1A9DA
      colorFocus      =   "frmRequest.frx":1AA04
      colorDisabled   =   "frmRequest.frx":1AA2E
      colorPressed    =   "frmRequest.frx":1AA58
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   7
      Left            =   9090
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4815
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1AA82
      textLT          =   "frmRequest.frx":1AAE4
      textCT          =   "frmRequest.frx":1AAFC
      textRT          =   "frmRequest.frx":1AB14
      textLM          =   "frmRequest.frx":1AB2C
      textRM          =   "frmRequest.frx":1AB44
      textLB          =   "frmRequest.frx":1AB5C
      textCB          =   "frmRequest.frx":1AB74
      textRB          =   "frmRequest.frx":1AB8C
      colorBack       =   "frmRequest.frx":1ABA4
      colorIntern     =   "frmRequest.frx":1ABCE
      colorMO         =   "frmRequest.frx":1ABF8
      colorFocus      =   "frmRequest.frx":1AC22
      colorDisabled   =   "frmRequest.frx":1AC4C
      colorPressed    =   "frmRequest.frx":1AC76
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   9
      Left            =   9090
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5955
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1ACA0
      textLT          =   "frmRequest.frx":1AD02
      textCT          =   "frmRequest.frx":1AD1A
      textRT          =   "frmRequest.frx":1AD32
      textLM          =   "frmRequest.frx":1AD4A
      textRM          =   "frmRequest.frx":1AD62
      textLB          =   "frmRequest.frx":1AD7A
      textCB          =   "frmRequest.frx":1AD92
      textRB          =   "frmRequest.frx":1ADAA
      colorBack       =   "frmRequest.frx":1ADC2
      colorIntern     =   "frmRequest.frx":1ADEC
      colorMO         =   "frmRequest.frx":1AE16
      colorFocus      =   "frmRequest.frx":1AE40
      colorDisabled   =   "frmRequest.frx":1AE6A
      colorPressed    =   "frmRequest.frx":1AE94
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   11
      Left            =   9090
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7095
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1AEBE
      textLT          =   "frmRequest.frx":1AF20
      textCT          =   "frmRequest.frx":1AF38
      textRT          =   "frmRequest.frx":1AF50
      textLM          =   "frmRequest.frx":1AF68
      textRM          =   "frmRequest.frx":1AF80
      textLB          =   "frmRequest.frx":1AF98
      textCB          =   "frmRequest.frx":1AFB0
      textRB          =   "frmRequest.frx":1AFC8
      colorBack       =   "frmRequest.frx":1AFE0
      colorIntern     =   "frmRequest.frx":1B00A
      colorMO         =   "frmRequest.frx":1B034
      colorFocus      =   "frmRequest.frx":1B05E
      colorDisabled   =   "frmRequest.frx":1B088
      colorPressed    =   "frmRequest.frx":1B0B2
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   13
      Left            =   9090
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8235
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1B0DC
      textLT          =   "frmRequest.frx":1B13E
      textCT          =   "frmRequest.frx":1B156
      textRT          =   "frmRequest.frx":1B16E
      textLM          =   "frmRequest.frx":1B186
      textRM          =   "frmRequest.frx":1B19E
      textLB          =   "frmRequest.frx":1B1B6
      textCB          =   "frmRequest.frx":1B1CE
      textRB          =   "frmRequest.frx":1B1E6
      colorBack       =   "frmRequest.frx":1B1FE
      colorIntern     =   "frmRequest.frx":1B228
      colorMO         =   "frmRequest.frx":1B252
      colorFocus      =   "frmRequest.frx":1B27C
      colorDisabled   =   "frmRequest.frx":1B2A6
      colorPressed    =   "frmRequest.frx":1B2D0
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1200
      Index           =   15
      Left            =   9090
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   9375
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2117
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1B2FA
      textLT          =   "frmRequest.frx":1B35C
      textCT          =   "frmRequest.frx":1B374
      textRT          =   "frmRequest.frx":1B38C
      textLM          =   "frmRequest.frx":1B3A4
      textRM          =   "frmRequest.frx":1B3BC
      textLB          =   "frmRequest.frx":1B3D4
      textCB          =   "frmRequest.frx":1B3EC
      textRB          =   "frmRequest.frx":1B404
      colorBack       =   "frmRequest.frx":1B41C
      colorIntern     =   "frmRequest.frx":1B446
      colorMO         =   "frmRequest.frx":1B470
      colorFocus      =   "frmRequest.frx":1B49A
      colorDisabled   =   "frmRequest.frx":1B4C4
      colorPressed    =   "frmRequest.frx":1B4EE
      Style           =   2
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   2
      Left            =   7620
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1B518
      textLT          =   "frmRequest.frx":1B57A
      textCT          =   "frmRequest.frx":1B592
      textRT          =   "frmRequest.frx":1B5AA
      textLM          =   "frmRequest.frx":1B5C2
      textRM          =   "frmRequest.frx":1B5DA
      textLB          =   "frmRequest.frx":1B5F2
      textCB          =   "frmRequest.frx":1B60A
      textRB          =   "frmRequest.frx":1B622
      colorBack       =   "frmRequest.frx":1B63A
      colorIntern     =   "frmRequest.frx":1B664
      colorMO         =   "frmRequest.frx":1B68E
      colorFocus      =   "frmRequest.frx":1B6B8
      colorDisabled   =   "frmRequest.frx":1B6E2
      colorPressed    =   "frmRequest.frx":1B70C
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   4
      Left            =   7620
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3675
      Visible         =   0   'False
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1B736
      textLT          =   "frmRequest.frx":1B798
      textCT          =   "frmRequest.frx":1B7B0
      textRT          =   "frmRequest.frx":1B7C8
      textLM          =   "frmRequest.frx":1B7E0
      textRM          =   "frmRequest.frx":1B7F8
      textLB          =   "frmRequest.frx":1B810
      textCB          =   "frmRequest.frx":1B828
      textRB          =   "frmRequest.frx":1B840
      colorBack       =   "frmRequest.frx":1B858
      colorIntern     =   "frmRequest.frx":1B882
      colorMO         =   "frmRequest.frx":1B8AC
      colorFocus      =   "frmRequest.frx":1B8D6
      colorDisabled   =   "frmRequest.frx":1B900
      colorPressed    =   "frmRequest.frx":1B92A
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   6
      Left            =   7620
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4815
      Visible         =   0   'False
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1B954
      textLT          =   "frmRequest.frx":1B9B6
      textCT          =   "frmRequest.frx":1B9CE
      textRT          =   "frmRequest.frx":1B9E6
      textLM          =   "frmRequest.frx":1B9FE
      textRM          =   "frmRequest.frx":1BA16
      textLB          =   "frmRequest.frx":1BA2E
      textCB          =   "frmRequest.frx":1BA46
      textRB          =   "frmRequest.frx":1BA5E
      colorBack       =   "frmRequest.frx":1BA76
      colorIntern     =   "frmRequest.frx":1BAA0
      colorMO         =   "frmRequest.frx":1BACA
      colorFocus      =   "frmRequest.frx":1BAF4
      colorDisabled   =   "frmRequest.frx":1BB1E
      colorPressed    =   "frmRequest.frx":1BB48
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   8
      Left            =   7620
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5955
      Visible         =   0   'False
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1BB72
      textLT          =   "frmRequest.frx":1BBD4
      textCT          =   "frmRequest.frx":1BBEC
      textRT          =   "frmRequest.frx":1BC04
      textLM          =   "frmRequest.frx":1BC1C
      textRM          =   "frmRequest.frx":1BC34
      textLB          =   "frmRequest.frx":1BC4C
      textCB          =   "frmRequest.frx":1BC64
      textRB          =   "frmRequest.frx":1BC7C
      colorBack       =   "frmRequest.frx":1BC94
      colorIntern     =   "frmRequest.frx":1BCBE
      colorMO         =   "frmRequest.frx":1BCE8
      colorFocus      =   "frmRequest.frx":1BD12
      colorDisabled   =   "frmRequest.frx":1BD3C
      colorPressed    =   "frmRequest.frx":1BD66
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   10
      Left            =   7620
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7095
      Visible         =   0   'False
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1BD90
      textLT          =   "frmRequest.frx":1BDF2
      textCT          =   "frmRequest.frx":1BE0A
      textRT          =   "frmRequest.frx":1BE22
      textLM          =   "frmRequest.frx":1BE3A
      textRM          =   "frmRequest.frx":1BE52
      textLB          =   "frmRequest.frx":1BE6A
      textCB          =   "frmRequest.frx":1BE82
      textRB          =   "frmRequest.frx":1BE9A
      colorBack       =   "frmRequest.frx":1BEB2
      colorIntern     =   "frmRequest.frx":1BEDC
      colorMO         =   "frmRequest.frx":1BF06
      colorFocus      =   "frmRequest.frx":1BF30
      colorDisabled   =   "frmRequest.frx":1BF5A
      colorPressed    =   "frmRequest.frx":1BF84
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   12
      Left            =   7620
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   8235
      Visible         =   0   'False
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1BFAE
      textLT          =   "frmRequest.frx":1C010
      textCT          =   "frmRequest.frx":1C028
      textRT          =   "frmRequest.frx":1C040
      textLM          =   "frmRequest.frx":1C058
      textRM          =   "frmRequest.frx":1C070
      textLB          =   "frmRequest.frx":1C088
      textCB          =   "frmRequest.frx":1C0A0
      textRB          =   "frmRequest.frx":1C0B8
      colorBack       =   "frmRequest.frx":1C0D0
      colorIntern     =   "frmRequest.frx":1C0FA
      colorMO         =   "frmRequest.frx":1C124
      colorFocus      =   "frmRequest.frx":1C14E
      colorDisabled   =   "frmRequest.frx":1C178
      colorPressed    =   "frmRequest.frx":1C1A2
      Style           =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1200
      Index           =   14
      Left            =   7620
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   9375
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2117
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1C1CC
      textLT          =   "frmRequest.frx":1C22E
      textCT          =   "frmRequest.frx":1C246
      textRT          =   "frmRequest.frx":1C25E
      textLM          =   "frmRequest.frx":1C276
      textRM          =   "frmRequest.frx":1C28E
      textLB          =   "frmRequest.frx":1C2A6
      textCB          =   "frmRequest.frx":1C2BE
      textRB          =   "frmRequest.frx":1C2D6
      colorBack       =   "frmRequest.frx":1C2EE
      colorIntern     =   "frmRequest.frx":1C318
      colorMO         =   "frmRequest.frx":1C342
      colorFocus      =   "frmRequest.frx":1C36C
      colorDisabled   =   "frmRequest.frx":1C396
      colorPressed    =   "frmRequest.frx":1C3C0
      Style           =   2
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   0
      Left            =   10830
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6090
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   1
      Left            =   12255
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6090
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   2
      Left            =   13680
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6090
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   3
      Left            =   10830
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   4
      Left            =   12255
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   5
      Left            =   13680
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   6
      Left            =   10830
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   8310
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   7
      Left            =   12255
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   8310
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   8
      Left            =   13680
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   8310
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   930
      Index           =   9
      Left            =   13080
      TabIndex        =   49
      Top             =   330
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1640
      Appearance      =   3
      BackColor       =   2720171
      Caption         =   "CL"
      CaptionOffsetY  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   10
      Left            =   12255
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   9420
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   11
      Left            =   13680
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   9420
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   930
      Index           =   1
      Left            =   14130
      TabIndex        =   52
      Top             =   330
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1640
      Appearance      =   3
      BackColor       =   2163158
      Caption         =   "X"
      CaptionOffsetY  =   2
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picHoldFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   690
      ScaleHeight     =   615
      ScaleWidth      =   825
      TabIndex        =   55
      Top             =   1140
      Width           =   825
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   910
      Index           =   3
      Left            =   7140
      TabIndex        =   57
      Top             =   330
      Width           =   795
      _Version        =   524298
      _ExtentX        =   1402
      _ExtentY        =   1605
      _StockProps     =   66
      Caption         =   "6"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   24.75
         Charset         =   2
         Weight          =   400
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
      CornerFactor    =   20
      BackColorContainer=   2323868
      ButtonRaiseFactor=   3
      SmoothEdges     =   1
      SpecialEffect   =   2
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      FlatPillowFactor=   3
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1C3EA
      textLT          =   "frmRequest.frx":1C44C
      textCT          =   "frmRequest.frx":1C464
      textRT          =   "frmRequest.frx":1C47C
      textLM          =   "frmRequest.frx":1C494
      textRM          =   "frmRequest.frx":1C4AC
      textLB          =   "frmRequest.frx":1C4C4
      textCB          =   "frmRequest.frx":1C4DC
      textRB          =   "frmRequest.frx":1C4F4
      colorBack       =   "frmRequest.frx":1C50C
      colorIntern     =   "frmRequest.frx":1C536
      colorMO         =   "frmRequest.frx":1C560
      colorFocus      =   "frmRequest.frx":1C58A
      colorDisabled   =   "frmRequest.frx":1C5B4
      colorPressed    =   "frmRequest.frx":1C5DE
      Orientation     =   3
      TextCaptionAlignment=   0
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   1
      Left            =   9090
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1410
      Visible         =   0   'False
      Width           =   1530
      _Version        =   524298
      _ExtentX        =   2699
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1C608
      textLT          =   "frmRequest.frx":1C66A
      textCT          =   "frmRequest.frx":1C682
      textRT          =   "frmRequest.frx":1C69A
      textLM          =   "frmRequest.frx":1C6B2
      textRM          =   "frmRequest.frx":1C6CA
      textLB          =   "frmRequest.frx":1C6E2
      textCB          =   "frmRequest.frx":1C6FA
      textRB          =   "frmRequest.frx":1C712
      colorBack       =   "frmRequest.frx":1C72A
      colorIntern     =   "frmRequest.frx":1C754
      colorMO         =   "frmRequest.frx":1C77E
      colorFocus      =   "frmRequest.frx":1C7A8
      colorDisabled   =   "frmRequest.frx":1C7D2
      colorPressed    =   "frmRequest.frx":1C7FC
      Style           =   2
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1140
      Index           =   0
      Left            =   7620
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   1410
      Visible         =   0   'False
      Width           =   1485
      _Version        =   524298
      _ExtentX        =   2619
      _ExtentY        =   2011
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   12
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1C826
      textLT          =   "frmRequest.frx":1C888
      textCT          =   "frmRequest.frx":1C8A0
      textRT          =   "frmRequest.frx":1C8B8
      textLM          =   "frmRequest.frx":1C8D0
      textRM          =   "frmRequest.frx":1C8E8
      textLB          =   "frmRequest.frx":1C900
      textCB          =   "frmRequest.frx":1C918
      textRB          =   "frmRequest.frx":1C930
      colorBack       =   "frmRequest.frx":1C948
      colorIntern     =   "frmRequest.frx":1C972
      colorMO         =   "frmRequest.frx":1C99C
      colorFocus      =   "frmRequest.frx":1C9C6
      colorDisabled   =   "frmRequest.frx":1C9F0
      colorPressed    =   "frmRequest.frx":1CA1A
      Style           =   2
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   12
      Left            =   10830
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   9420
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1905
      Appearance      =   3
      BackColor       =   6208225
      Caption         =   "X"
      CaptionOffsetY  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin BTNENHLib4.BtnEnh cmdUser 
      Height          =   910
      Left            =   4570
      TabIndex        =   56
      Top             =   330
      Width           =   2625
      _Version        =   524298
      _ExtentX        =   4630
      _ExtentY        =   1605
      _StockProps     =   66
      Caption         =   "<Request From>"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
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
      CornerFactor    =   20
      BackColorContainer=   2323868
      ButtonRaiseFactor=   3
      SmoothEdges     =   1
      SpecialEffect   =   2
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      FlatPillowFactor=   3
      UserData        =   0.1
      textCaption     =   "frmRequest.frx":1CA44
      textLT          =   "frmRequest.frx":1CAC0
      textCT          =   "frmRequest.frx":1CAD8
      textRT          =   "frmRequest.frx":1CAF0
      textLM          =   "frmRequest.frx":1CB08
      textRM          =   "frmRequest.frx":1CB20
      textLB          =   "frmRequest.frx":1CB38
      textCB          =   "frmRequest.frx":1CB50
      textRB          =   "frmRequest.frx":1CB68
      colorBack       =   "frmRequest.frx":1CB80
      colorIntern     =   "frmRequest.frx":1CBAA
      colorMO         =   "frmRequest.frx":1CBD4
      colorFocus      =   "frmRequest.frx":1CBFE
      colorDisabled   =   "frmRequest.frx":1CC28
      colorPressed    =   "frmRequest.frx":1CC52
      Orientation     =   1
      TextCaptionAlignment=   0
      HollowFrame     =   -1  'True
   End
   Begin VB.Label lblWeight_Empty 
      Height          =   285
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblWeight_Full 
      Height          =   285
      Left            =   3060
      TabIndex        =   70
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label QAnswer 
      Height          =   345
      Left            =   1680
      TabIndex        =   69
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSForms.Label lblDepart 
      Height          =   285
      Left            =   6420
      TabIndex        =   60
      Top             =   10860
      Width           =   3795
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "6694;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label lblDetail 
      Height          =   675
      Left            =   8190
      TabIndex        =   53
      Top             =   480
      Width           =   4605
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "8123;1191"
      FontName        =   "Arial Narrow"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblUser 
      Height          =   285
      Left            =   10290
      TabIndex        =   38
      Top             =   10860
      Width           =   4365
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7699;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   480
      TabIndex        =   39
      Top             =   10860
      Width           =   4005
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7064;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Image newBack 
      Height          =   1815
      Left            =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "2725;3201"
      Picture         =   "frmRequest.frx":1CC7C
   End
End
Attribute VB_Name = "frmRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LoadPlu(Dept_No)
    Screen.MousePointer = 11
    grdPlu.Rows = 0
    cmdPlu(0).Caption = ""
    cmdPlu(0).Picture = ""
    cmdPlu(0).ToolTipText = ""
    ActiveReadServer "SELECT Description,Short_Description,Unit_Size,Unit_of_Measure, Product_Code FROM  Products where Department_No= '" & Dept_No & " ' and Stock_Item=1 order by Description"
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdPlu.Rows = grdPlu.Rows + 1
        If rs.Fields("Unit_Size") = 0 Then
            unit_Size = ""
        Else
            unit_Size = rs.Fields("Unit_Size")
        End If
        If rs.Fields("Short_Description") & "" = "" Then
            Description = rs.Fields("Description")
        Else
            Description = rs.Fields("Short_Description")
        End If
        If i < 17 And Not rs.EOF Then
            cmdPlu(i).Caption = Replace(Description, "&", "&&") & " " & unit_Size & rs.Fields("Unit_of_Measure")
            cmdPlu(i).Tag = rs.Fields("Product_Code")
            cmdPlu(i).TextDescrCB.Text = ""
            cmdPlu(i).ToolTipText = " Product Code: " & rs.Fields("Product_Code") & " "
            If grdMain.FindRow(rs.Fields("Product_Code"), 0, 2) = -1 Then
                cmdPlu(i).TextDescrCB.Text = "Not Requested"
                cmdPlu(i).TextDescrCB.OffsetY = -8
                cmdPlu(i).TextDescrCB.ColorNormal = &HC0&
            Else
                cmdPlu(i).TextDescrCB.Text = ""
            End If
            cmdPlu(i).Value = 0
            If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
            grdPlu.Row = grdPlu.Rows - 1
            grdPlu.TextMatrix(grdPlu.Rows - 1, 0) = Replace(Description, "&", "&&") & " " & unit_Size & rs.Fields("Unit_of_Measure")
            grdPlu.TextMatrix(grdPlu.Rows - 1, 1) = rs.Fields("Product_Code")
        Else
            If b = 0 Then
                grdPlu.TextMatrix(grdPlu.Rows - 1, 1) = "ßArrow"
                grdPlu.Rows = grdPlu.Rows + 1
                If i = 17 Then
                    cmdPlu(17).Value = 0
                    cmdPlu(17).Caption = ""
                    cmdPlu(17).TextDescrCB.Text = ""
                    cmdPlu(17).ToolTipText = ""
                    cmdPlu(17).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdPlu(17).Visible = False Then cmdPlu(17).Visible = True
                End If
            End If
            b = b + 1
            grdPlu.TextMatrix(grdPlu.Rows - 1, 0) = Replace(Description, "&", "&&") & " " & unit_Size & rs.Fields("Unit_of_Measure")
            grdPlu.TextMatrix(grdPlu.Rows - 1, 1) = rs.Fields("Product_Code")
            If b = 16 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdPlu.Count - 1
       cmdPlu(b).Caption = "1"
       cmdPlu(b).Tag = ""
       cmdPlu(b).ToolTipText = ""
       cmdPlu(b).TextDescrCB.Text = ""
       If cmdPlu(b).Visible = True Then cmdPlu(b).Visible = False
    Next b
    Screen.MousePointer = 0
End Sub
Private Sub BtnEnh1_Click(Index As Integer)
    If picSlip.Visible = True Then Exit Sub
    If Timer2.Enabled = True Then Exit Sub
    Select Case grdLoc.Visible
         Case True
            grdLoc.Visible = False
         Case False
            grdLoc.Rows = 0
            grdLoc.Visible = True
            ActiveReadServer "Select * from Locations where Stock_Take = 0 order by Location_No"
            While Not rs.EOF
                grdLoc.Rows = grdLoc.Rows + 1
                grdLoc.TextMatrix(grdLoc.Rows - 1, 0) = rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            If grdLoc.Rows > 0 Then
                grdLoc.Height = 650 * grdLoc.Rows - 1
                If grdLoc.Rows > 0 Then grdLoc.Row = 0
                grdLoc.SetFocus
            Else
                grdLoc.Visible = False
            End If
    End Select
    Me.SetFocus
End Sub
Private Sub BtnEnh1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub
Private Sub cmdArrow_Click(Index As Integer)
    Select Case Index
        Case 0
            If grdRem.Row <> 1 Then
                grdRem.Row = grdRem.Row - 1
                grdMain.Row = grdMain.Row - 1
            End If
        Case 1
            If grdRem.Row <> grdRem.Rows - 1 Then
                grdRem.Row = grdRem.Row + 1
                grdMain.Row = grdMain.Row + 1
            End If
    End Select
    grdRem.ShowCell grdRem.Row, 0
    grdMain.ShowCell grdMain.Row, 0
End Sub
Private Sub cmdArrow_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    scrolTimer.Tag = Index
    scrolTimer.Enabled = True
End Sub


Private Sub cmdDeptStrip_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub

Private Sub cmdErr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub
Private Sub cmdFancy_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
    picHoldFocus.SetFocus
End Sub
Private Sub cmdInput_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
    Me.SetFocus
End Sub
Private Sub cmdLogOff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub
Private Sub cmdPlu_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
    picHoldFocus.SetFocus
End Sub

Private Sub cmdRemove_Click()
    If grdRem.Rows = 1 Then Exit Sub
    ActiveUpdateServer "Delete from Stock_Request_Listing where Line_No = " & grdRem.TextMatrix(grdRem.Row, 3)
    grdRem.RemoveItem (grdRem.Row)
    grdMain.RemoveItem (grdMain.Row)
    ActiveReadServer "Select * from Departments_Stock where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & " and Department_No = '" & Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) & "'"
    If rs.RecordCount > 0 Then
        cmdDeptStrip(Val(lblDepart.Tag)).TextDescrCB.Text = rs.Fields("Product_Count") & " Requested of " & rs.Fields("Product_List")
        For i = 0 To grdDept.Rows - 1
            If grdDept.TextMatrix(i, 1) = Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) Then
                grdDept.TextMatrix(i, 3) = rs.Fields("Product_Count") & " Requested of " & rs.Fields("Product_List")
            End If
        Next i
    End If
    rs.Close
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Right(lblDetail.Caption, 2) = " " & Chr(215) Then
            KeyCode = 0
            Exit Sub
        End If
        KeyCode = 0
        ActiveReadServer "Select * from Products where Product_code = '" & Trim(Mid(lblDetail.Caption, InStr(lblDetail.Caption, Chr(215)) + 1)) & "' and Stock_Item =1"
        If rs.RecordCount > 0 Then
            Request_No = 0
            Qty_Counted = lblDetail.Caption
            If InStr(lblDetail.Caption, Chr(215)) <> 0 Then
                Qty_Counted = Val(Trim(Mid(lblDetail.Caption, 1, InStr(lblDetail.Caption, Chr(215)) - 1)))
                Product_Code = Trim(Mid(lblDetail.Caption, InStr(lblDetail.Caption, Chr(215)) + 1))
            Else
                Qty_Counted = 1
                Product_Code = lblDetail.Caption
            End If
            Unit_of_Measure = "each"
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = rs.Fields("Short_Description")
            grdMain.TextMatrix(grdMain.Row, 1) = Qty_Counted
            grdMain.TextMatrix(grdMain.Row, 2) = Product_Code
            lblDetail.Caption = "Requested: " & Qty_Counted & " x " & rs.Fields("Short_Description")
            
            ActiveUpdateServer "INSERT INTO Stock_Request_Listing (Date_Time,  Workstation_No, User_No, Location_No,Product_Code,Department_No,Qty_Counted,Unit_of_Measure) values " & _
            "(Getdate()," & Workstation_No & "," & UserRecord.User_Number & "," & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & ",'" & Product_Code & "','" & rs.Fields("Department_No") & "'," & Qty_Counted & ",'" & Unit_of_Measure & "')"
            
            ActiveReadServer1 "Select max(Line_No) as Line_No from Stock_Request_Listing"
            grdMain.TextMatrix(grdMain.Row, 3) = rs1.Fields("Line_No")
            rs1.Close
            ActiveReadServer1 "Select * from Departments_Stock where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & " and Department_No = '" & rs.Fields("Department_No") & "'"
            If rs1.RecordCount > 0 Then
                cmdDeptStrip(Val(lblDepart.Tag)).TextDescrCB.Text = rs1.Fields("Product_Count") & " Requested of " & rs1.Fields("Product_List")
                For i = 0 To grdDept.Rows - 1
                    If grdDept.TextMatrix(i, 1) = rs.Fields("Department_No") Then
                        grdDept.TextMatrix(i, 3) = rs1.Fields("Product_Count") & " Requested of " & rs1.Fields("Product_List")
                    End If
                Next i
            End If
            rs1.Close
            grdMain.ShowCell grdMain.Rows - 1, 0
            lblDetail.Font.Size = 16
            Me.SetFocus
        Else
            cmdErr.Caption = "Unknown Product"
            Timer2.Enabled = True
            cmdErr.Visible = True
            lblDetail.Caption = ""
            Exit Sub
        End If
        rs.Close
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If cmdUser.Caption = "<Request From>" Then
        cmdErr.Caption = "Select a Location First"
        Timer2.Enabled = True
        cmdErr.Visible = True
        KeyAscii = 0
        Exit Sub
    End If
    Select Case KeyAscii
        Case 42
            If InStr(lblDetail.Caption, Chr(215)) <> 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            lblDetail.Caption = lblDetail.Caption & " " & Chr(215) & " "
            lblDetail.Font.Size = 26
            If InStr(lblDetail.Caption, Chr(215)) <> 0 Then
                KeyAscii = 0
                Exit Sub
            End If
        Case 13
            KeyAscii = 0
        Case 8
            lblDetail.Caption = ""
        Case 97 To 122
            lblDetail.Font.Size = 26
            KeyAscii = KeyAscii - 32
            lblDetail.Caption = lblDetail.Caption & Chr(KeyAscii)
            KeyAscii = 0
        Case 65 To 90
            lblDetail.Font.Size = 26
            lblDetail.Caption = lblDetail.Caption & Chr(KeyAscii)
            KeyAscii = 0
        Case 46, 48 To 57
            lblDetail.Font.Size = 26
            If lblDetail.Caption <> "" Then
                If Asc(Mid(lblDetail.Caption, 1, 1)) < 48 Or Asc(Mid(lblDetail.Caption, 1, 1)) > 57 Then
                    If lblDetail.Caption = "." Then
                       lblDetail.Caption = lblDetail.Caption & Chr(KeyAscii)
                    Else
                        If InStr(lblDetail.Caption, ".") = 0 Then
                            lblDetail.Caption = Chr(KeyAscii)
                        Else
                            If Left(lblDetail.Caption, 9) = "Requested" Then
                                lblDetail.Caption = Chr(KeyAscii)
                            Else
                                lblDetail.Caption = lblDetail.Caption & Chr(KeyAscii)
                            End If
                        End If
                    End If
                Else
                    lblDetail.Caption = lblDetail.Caption & Chr(KeyAscii)
                End If
            Else
                lblDetail.Caption = lblDetail.Caption & Chr(KeyAscii)
            End If
            KeyAscii = 0
    End Select
End Sub
Private Sub grdMain_Click()
    On Error Resume Next
    If grdRem.Row <> 0 Then
        grdRem.Row = grdMain.Row
        grdRem.ShowCell grdRem.Row, 0
    End If
    On Error GoTo 0
End Sub

Private Sub grdRem_Click()
    grdMain.Row = grdRem.Row
    grdMain.ShowCell grdMain.Row, 0
End Sub

Private Sub scrolTimer_Timer()
    scrolTimer.Interval = 50
    Select Case scrolTimer.Tag
        Case "0"
            If grdRem.Row <> 1 Then
                grdRem.Row = grdRem.Row - 1
                grdMain.Row = grdMain.Row - 1
            End If
        Case "1"
            If grdRem.Row <> grdRem.Rows - 1 Then
                grdMain.Row = grdMain.Row + 1
                grdRem.Row = grdRem.Row + 1
            End If
    End Select
    grdRem.ShowCell grdRem.Row, 0
    grdMain.ShowCell grdMain.Row, 0
End Sub
Private Sub cmdClose_Click()
    picSlip.Visible = False
End Sub
Private Sub cmdDeptStrip_Click(Index As Integer)
    Me.SetFocus
    DoEvents
    If picSlip.Visible = True Then
        cmdDeptStrip(Index).Value = 0
        Exit Sub
    End If
    If cmdDeptStrip(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        cmdDeptStrip(Index).Value = 0
        cmdDeptStrip(Index).TextDescrCB.ColorNormal = vbYellow
        DoEvents
        grdDept.Row = grdDept.Row + 1
        For i = 0 To 15
            If grdDept.TextMatrix(grdDept.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdDeptStrip(i).Caption = ""
                    cmdDeptStrip(i).Value = 0
                    cmdDeptStrip(i).TextDescrCB.Text = ""
                    cmdDeptStrip(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                Else
                    cmdDeptStrip(i).Caption = ""
                    cmdDeptStrip(i).Value = 0
                    cmdDeptStrip(i).TextDescrCB.Text = ""
                    cmdDeptStrip(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                    grdDept.Row = grdDept.Row - 1
                    Exit For
                End If
            Else
                cmdDeptStrip(i).Caption = grdDept.TextMatrix(grdDept.Row, 0)
                cmdDeptStrip(i).Tag = grdDept.TextMatrix(grdDept.Row, 1)
                cmdDeptStrip(i).TextDescrCB.Text = grdDept.TextMatrix(grdDept.Row, 3)
            End If
            If grdDept.Row = grdDept.Rows - 1 Then Exit For
            grdDept.Row = grdDept.Row + 1
        Next i
        For b = i + 1 To cmdDeptStrip.Count - 1
            cmdDeptStrip(b).Caption = "1"
            cmdDeptStrip(b).Tag = ""
            cmdDeptStrip(b).TextDescrCB.Text = "1"
            cmdDeptStrip(b).Visible = False
        Next b
    End If
    If cmdDeptStrip(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdDeptStrip(Index).Value = 0
        cmdDeptStrip(Index).TextDescrCB.ColorNormal = vbYellow
        DoEvents
        cmdDeptStrip(0).Picture = ""
        While grdDept.TextMatrix(grdDept.Row, 1) <> "ßArrow"
            grdDept.Row = grdDept.Row - 1
        Wend
        grdDept.Row = grdDept.Row - 15
        For i = 0 To 15
            If grdDept.TextMatrix(grdDept.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdDeptStrip(i).Caption = ""
                    cmdDeptStrip(i).Value = 0
                    cmdDeptStrip(i).TextDescrCB.Text = ""
                    cmdDeptStrip(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                Else
                    cmdDeptStrip(i).Caption = ""
                    cmdDeptStrip(i).Value = 0
                    cmdDeptStrip(i).TextDescrCB.Text = ""
                    cmdDeptStrip(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
                    grdDept.Row = grdDept.Row - 1
                    Exit For
                End If
            Else
                cmdDeptStrip(i).Caption = grdDept.TextMatrix(grdDept.Row, 0)
                cmdDeptStrip(i).TextDescrCB.Text = grdDept.TextMatrix(grdDept.Row, 3)
                cmdDeptStrip(i).Tag = grdDept.TextMatrix(grdDept.Row, 1)
                If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
            End If
            If grdDept.Row = grdDept.Rows - 1 Then Exit For
            grdDept.Row = grdDept.Row + 1
        Next i
        For b = i + 1 To cmdDeptStrip.Count - 1
            cmdDeptStrip(b).Caption = "1"
            cmdDeptStrip(b).TextDescrCB.Text = "1"
            cmdDeptStrip(b).Tag = ""
            cmdDeptStrip(b).Visible = False
        Next b
    End If
    Me.SetFocus
End Sub

Private Sub cmdDeptStrip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    For i = 0 To 15
        If Index <> i Then
            cmdDeptStrip(i).TextDescrCB.ColorNormal = vbYellow
        Else
            cmdDeptStrip(i).TextDescrCB.ColorNormal = &HC0&
        End If
    Next i
    If cmdDeptStrip(Index).Picture <> App.Path & "\icons\downArr.bmp" And cmdDeptStrip(Index).Picture <> App.Path & "\icons\upArr.bmp" Then
        lblDepart.Caption = "Counting: " & cmdDeptStrip(Index).Tag & " > " & Replace(cmdDeptStrip(Index).Caption, "&&", "&")
        lblDepart.Tag = Index
        LoadPlu cmdDeptStrip(Index).Tag
    End If
    DoEvents
End Sub

Private Sub cmdErr_Click()
    cmdErr.Caption = ""
    Timer2.Enabled = False
    cmdErr.Visible = False
End Sub
Private Sub Finalize_Take()
'    If grdMain.Rows = 1 Then Exit Sub
'    QAnswer.Caption = ""
'    Load frmQuestion
'    frmQuestion.Tag = "Take"
'    frmStockTake.Tag = "Not Now"
'    frmQuestion.lblCap = "Are you sure you want Finalize this Stock Take?"
'    frmQuestion.Show vbModal
'    Select Case QAnswer.Caption
'        Case "Yes"
'            Screen.MousePointer = 11
'            For i = 0 To cmdDeptStrip.Count - 1
'                cmdDeptStrip(i).Value = 0
'                cmdDeptStrip(i).TextDescrCB.ColorNormal = vbYellow
'            Next i
'            Load_List 0, Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
'            Merge_Take
'            ActiveReadServer1 "Select isnull(max(Request_No),0)+1 as Request_No from Stock_Take_Journal"
'            Request_No = rs1.Fields("Request_No")
'            rs1.Close
'            For i = 1 To grdMerge.Rows - 1
'                unit_Size = ""
'                Unit_of_Measure = ""
'                If InStr(grdMerge.TextMatrix(i, 2), "&") <> 0 Then
'                    Full_Units = Val(Mid(grdMerge.TextMatrix(i, 2), 1, InStr(grdMerge.TextMatrix(i, 2), "&") - 1))
'                    grdMerge.TextMatrix(i, 2) = Trim(Mid(grdMerge.TextMatrix(i, 2), InStr(grdMerge.TextMatrix(i, 2), "&") + 1))
'                Else
'                    Full_Units = 0
'                End If
'                For b = 1 To Len(grdMerge.TextMatrix(i, 2))
'                    If Asc(Mid(grdMerge.TextMatrix(i, 2), b, 1)) < 46 Or Asc(Mid(grdMerge.TextMatrix(i, 2), b, 1)) > 57 Then
'                        Unit_of_Measure = Trim(Mid(grdMerge.TextMatrix(i, 2), b))
'                        Exit For
'                    Else
'                        unit_Size = unit_Size & Mid(grdMerge.TextMatrix(i, 2), b, 1)
'                    End If
'                Next b
'                Select Case Unit_of_Measure
'                    Case "(Tots)"
'                       Bottle_Size = ""
'                       Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
'                        For b = 1 To Len(Bottle)
'                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
'                                Exit For
'                            Else
'                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
'                            End If
'                        Next b
'                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
'                        If Bottle_Size = "1000" Then
'                            unit_Size = Round((unit_Size * 40) / Val(Bottle_Size), 4)
'                        Else
'                            unit_Size = Round((unit_Size * 25) / Val(Bottle_Size), 4)
'                        End If
'                        grdMerge.TextMatrix(i, 2) = unit_Size
'                    Case "ml"
'                        Bottle_Size = ""
'                        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
'                        For b = 1 To Len(Bottle)
'                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
'                                Exit For
'                            Else
'                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
'                            End If
'                        Next b
'                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
'                        unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
'                        grdMerge.TextMatrix(i, 2) = unit_Size
'                    Case "kg"
'                        Bottle_Size = ""
'                        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
'                        For b = 1 To Len(Bottle)
'                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
'                                Exit For
'                            Else
'                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
'                            End If
'                        Next b
'                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
'                        unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
'                        grdMerge.TextMatrix(i, 2) = unit_Size
'                    Case "g"
'                        Bottle_Size = ""
'                        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
'                        For b = 1 To Len(Bottle)
'                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
'                                Exit For
'                            Else
'                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
'                            End If
'                        Next b
'                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
'                        unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
'                        grdMerge.TextMatrix(i, 2) = unit_Size
'                    Case "lt"
'                        Bottle_Size = ""
'                        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
'                        For b = 1 To Len(Bottle)
'                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
'                                Exit For
'                            Else
'                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
'                            End If
'                        Next b
'                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
'                        unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
'                        grdMerge.TextMatrix(i, 2) = unit_Size
'                End Select
'                Qty_Count = grdMerge.TextMatrix(i, 2) + Full_Units
'                ActiveReadServer "Select Stock_on_Hand from Quantities where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "' and Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
'                If rs.RecordCount > 0 Then
'                    Stock_on_Hand = rs.Fields("Stock_on_Hand")
'                    ActiveUpdateServer "Update Quantities Set Stock_on_Hand = " & Qty_Count & " where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "' and Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
'                Else
'                    Stock_on_Hand = 0
'                    ActiveUpdateServer "INSERT INTO Quantities (Product_Code,Location_No,Stock_on_Hand) values ('" & grdMerge.TextMatrix(i, 4) & "'," & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & "," & Qty_Count & ")"
'                End If
'                rs.Close
'                DoEvents
'                ActiveReadServer "Select Ave_Cost from Products where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "'"
'                If rs.RecordCount > 0 Then
'                    Ave_Cost = rs.Fields("Ave_Cost")
'                Else
'                    Ave_Cost = 0
'                End If
'                DoEvents
'                ActiveUpdateServer "INSERT INTO Stock_Take_Journal(Request_No, Date_Time, Workstation_No, User_No, Location_No, Product_Code, Qty_on_Hand, Qty_Counted,Ave_Cost)" & _
'                " VALUES(" & Request_No & ",Getdate()," & Workstation_No & "," & UserRecord.User_Number & "," & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & ",'" & grdMerge.TextMatrix(i, 4) & "'," & Stock_on_Hand & ",'" & Qty_Count & "'," & Ave_Cost & ")"
'                DoEvents
'            Next i
'            ActiveUpdateServer "Delete from Stock_Request_Listing where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
'            DoEvents
'            frmStockTake.Tag = ""
'            Screen.MousePointer = 11
'            Form_Activate
'            Screen.MousePointer = 0
'            MsgBox "Your Stock Take was Finalized as Request_No = " & Request_No, vbInformation, "HeroPOS"
'    End Select
'    Screen.MousePointer = 0
End Sub
Private Sub cmdFancy_Click(Index As Integer)
    If picSlip.Visible = True Then Exit Sub
    Select Case cmdFancy(Index).Caption
        Case "Finalize Take"
            If cmdUser.Caption = "<Request From>" Then
                cmdErr.Caption = "Select a Location First"
                Timer2.Enabled = True
                cmdErr.Visible = True
                Exit Sub
            End If
            If picMerge.Visible = False Then
                cmdErr.Caption = "Merge First"
                Timer2.Enabled = True
                cmdErr.Visible = True
                Exit Sub
            End If
            If grdMerge.Rows = 1 Then
                cmdErr.Caption = "Nothing to Finalize"
                Timer2.Enabled = True
                cmdErr.Visible = True
                Exit Sub
            End If
            Finalize_Take
    End Select
    picHoldFocus.SetFocus
End Sub
Private Sub cmdInput_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If picSlip.Visible = True Then Exit Sub
    If cmdUser.Caption = "<Request From>" Then
        cmdErr.Caption = "Select a Location First"
        Timer2.Enabled = True
        cmdErr.Visible = True
        Exit Sub
    End If
    Select Case cmdInput(Index).Caption
        Case "X"
            If InStr(lblDetail.Caption, Chr(215)) <> 0 Then Exit Sub
            lblDetail.Caption = lblDetail.Caption & " " & Chr(215) & " "
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
            lblDetail.Font.Size = 26
            If lblDetail.Caption <> "" Then
                If Asc(Mid(lblDetail.Caption, 1, 1)) < 48 Or Asc(Mid(lblDetail.Caption, 1, 1)) > 57 Then
                    If lblDetail.Caption = "." Then
                       lblDetail.Caption = lblDetail.Caption & cmdInput(Index).Caption
                    Else
                        If InStr(lblDetail.Caption, ".") = 0 Then
                            lblDetail.Caption = cmdInput(Index).Caption
                        Else
                            If Left(lblDetail.Caption, 9) = "Requested" Then
                                lblDetail.Caption = cmdInput(Index).Caption
                            Else
                                lblDetail.Caption = lblDetail.Caption & cmdInput(Index).Caption
                            End If
                        End If
                    End If
                Else
                    lblDetail.Caption = lblDetail.Caption & cmdInput(Index).Caption
                End If
            Else
                lblDetail.Caption = lblDetail.Caption & cmdInput(Index).Caption
            End If
        Case "."
            If lblDetail.Caption <> "" Then
                If Asc(Mid(lblDetail.Caption, 1, 1)) < 45 Or Asc(Mid(lblDetail.Caption, 1, 1)) > 57 Then
                    lblDetail.Caption = ""
                End If
            End If
            lblDetail.Font.Size = 26
            If InStr(lblDetail.Caption, ".") <> 0 Then
                Exit Sub
            End If
            If lblDetail.Caption = "" Then
                lblDetail.Caption = cmdInput(Index).Caption
                Exit Sub
            End If
            If Asc(Mid(lblDetail.Caption, 1, 1)) < 48 Or Asc(Mid(lblDetail.Caption, 1, 1)) > 57 Then
                lblDetail.Caption = cmdInput(Index).Caption
            Else
                lblDetail.Caption = lblDetail.Caption & cmdInput(Index).Caption
            End If
        Case "CL"
            lblDetail.Font.Size = 26
            cmdErr.Caption = ""
            Timer2.Enabled = False
            cmdErr.Visible = False
            lblDetail.Caption = ""
    End Select
    Me.SetFocus
End Sub
Private Sub cmdKey_Click(Index As Integer)
    grdLoc.Visible = False
    Unload Me
End Sub
Private Sub cmdLogOff_Click()
    DoEvents
    Select Case UserRecord.uType
        Case 2, 3, 4, 8
            cmdErr.Caption = "Access Denied"
            Timer2.Enabled = True
            cmdErr.Visible = True
            Exit Sub
    End Select
    If cmdUser.Caption = "<Request From>" Then
        cmdErr.Caption = "Nothing to Clear"
        Timer2.Enabled = True
        cmdErr.Visible = True
        Exit Sub
    End If
    QAnswer.Caption = ""
    Load frmQuestion
    frmQuestion.Tag = "Take"
    frmStockTake.Tag = "Not Now"
    frmQuestion.lblCap = "Are you sure you want Clear this Stock Take?"
    frmQuestion.Show vbModal
    Select Case QAnswer.Caption
        Case "Yes"
            ActiveUpdateServer "Delete from Stock_Request_Listing where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
            QAnswer.Caption = ""
            lblDetail.Caption = "Select a Location to Request From"
            For i = 0 To cmdDeptStrip.Count - 1
                cmdDeptStrip(i).Visible = False
            Next i
            cmdUser.Caption = "<Request From>"
            grdMain.Rows = 1
            For i = 0 To cmdPlu.Count - 1
                cmdPlu(i).Visible = False
            Next i
            picMerge.Visible = False
        Case "No"
            DoEvents
    End Select
End Sub

Private Sub cmdPlu_Click(Index As Integer)
    If picSlip.Visible = True Then
        cmdPlu(Index).Value = 0
        Exit Sub
    End If
    DoEvents
    If cmdPlu(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        grdPlu.Row = grdPlu.Row + 1
        For i = 0 To 17
            If grdPlu.TextMatrix(grdPlu.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdPlu(i).Caption = ""
                    cmdPlu(i).ToolTipText = ""
                    cmdPlu(i).TextDescrCB.Text = ""
                    cmdPlu(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                Else
                    cmdPlu(i).Caption = ""
                    cmdPlu(i).ToolTipText = ""
                    cmdPlu(i).TextDescrCB.Text = ""
                    cmdPlu(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                    grdPlu.Row = grdPlu.Row - 1
                    Exit For
                End If
            Else
                cmdPlu(i).Caption = grdPlu.TextMatrix(grdPlu.Row, 0)
                cmdPlu(i).Tag = grdPlu.TextMatrix(grdPlu.Row, 1)
                cmdPlu(i).ToolTipText = " Product Code: " & grdPlu.TextMatrix(grdPlu.Row, 1) & " "
                If grdMain.FindRow(grdPlu.TextMatrix(grdPlu.Row, 1), 0, 2) = -1 Then
                    cmdPlu(i).TextDescrCB.Text = "Not Requested"
                    cmdPlu(i).TextDescrCB.OffsetY = -8
                    cmdPlu(i).TextDescrCB.ColorNormal = &HC0&
                Else
                    cmdPlu(i).TextDescrCB.Text = ""
                End If
            End If
            If grdPlu.Row = grdPlu.Rows - 1 Then Exit For
            grdPlu.Row = grdPlu.Row + 1
        Next i
        For b = i + 1 To cmdPlu.Count - 1
            cmdPlu(b).Caption = "1"
            cmdPlu(b).ToolTipText = ""
            cmdPlu(b).Tag = ""
            cmdPlu(b).Visible = False
        Next b
        cmdPlu(Index).Value = 0
    End If
    If cmdPlu(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdPlu(0).Picture = ""
        While grdPlu.TextMatrix(grdPlu.Row, 1) <> "ßArrow"
            grdPlu.Row = grdPlu.Row - 1
        Wend
        grdPlu.Row = grdPlu.Row - 17
        For i = 0 To 17
            If grdPlu.TextMatrix(grdPlu.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdPlu(i).Caption = ""
                    cmdPlu(i).ToolTipText = ""
                    cmdPlu(i).TextDescrCB.Text = ""
                    cmdPlu(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                Else
                    cmdPlu(i).Caption = ""
                     cmdPlu(i).ToolTipText = ""
                     cmdPlu(i).TextDescrCB.Text = ""
                    cmdPlu(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                    grdPlu.Row = grdPlu.Row - 1
                    Exit For
                End If
            Else
                cmdPlu(i).Caption = grdPlu.TextMatrix(grdPlu.Row, 0)
                cmdPlu(i).Tag = grdPlu.TextMatrix(grdPlu.Row, 1)
                cmdPlu(i).ToolTipText = " Product Code: " & grdPlu.TextMatrix(grdPlu.Row, 1) & " "
                If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                If grdMain.FindRow(grdPlu.TextMatrix(grdPlu.Row, 1), 0, 2) = -1 Then
                    cmdPlu(i).TextDescrCB.Text = "Not Requested"
                    cmdPlu(i).TextDescrCB.OffsetY = -8
                    cmdPlu(i).TextDescrCB.ColorNormal = &HC0&
                Else
                    cmdPlu(i).TextDescrCB.Text = ""
                End If
            End If
            If grdPlu.Row = grdPlu.Rows - 1 Then Exit For
            grdPlu.Row = grdPlu.Row + 1
        Next i
        For b = i + 1 To cmdPlu.Count - 1
            cmdPlu(b).Caption = "1"
            cmdPlu(b).Tag = ""
            cmdPlu(b).ToolTipText = ""
            cmdPlu(b).Visible = False
        Next b
        cmdPlu(Index).Value = 0
    End If
End Sub
Private Sub cmdPlu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdPlu(Index).Picture <> App.Path & "\icons\downArr.bmp" And cmdPlu(Index).Picture <> App.Path & "\icons\upArr.bmp" Then
        If Right(lblDetail.Caption, 3) = " " & Chr(215) & " " Then
            lblDetail.Caption = Left(lblDetail.Caption, Len(lblDetail.Caption) - 3)
        End If
        DoEvents
        lblDetail.Font.Size = 16
        If Trim(lblDetail.Caption) = "" Then lblDetail.Caption = "1"
        If Asc(Mid(lblDetail.Caption, 1, 1)) < 48 Or Asc(Mid(lblDetail.Caption, 1, 1)) > 57 Then
            If InStr(lblDetail.Caption, ".") = 0 Then
                lblDetail.Caption = "1"
            End If
        End If
        Request_No = 0
        Qty_Counted = lblDetail.Caption
        
        Unit_of_Measure = "each"
        grdMain.Rows = grdMain.Rows + 1
        grdMain.Row = grdMain.Rows - 1
        grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Index).Caption
        grdMain.TextMatrix(grdMain.Row, 1) = lblDetail.Caption
        grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Index).Tag
        lblDetail.Caption = "Requested: " & lblDetail.Caption & " x " & cmdPlu(Index).Caption
        cmdPlu(Index).TextDescrCB.Text = ""
        ActiveReadServer "Select * from Departments_Stock where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & " and Department_No = '" & Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) & "'"
        If rs.RecordCount > 0 Then
            cmdDeptStrip(Val(lblDepart.Tag)).TextDescrCB.Text = rs.Fields("Product_Count") & " Requested of " & rs.Fields("Product_List")
            For i = 0 To grdDept.Rows - 1
                If grdDept.TextMatrix(i, 1) = Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) Then
                    grdDept.TextMatrix(i, 3) = rs.Fields("Product_Count") & " Requested of " & rs.Fields("Product_List")
                End If
            Next i
        End If
        rs.Close
        grdMain.ShowCell grdMain.Rows - 1, 0
        Me.SetFocus
    End If
End Sub

Private Sub cmdUser_Click()
    If picSlip.Visible = True Then Exit Sub
    If Timer2.Enabled = True Then Exit Sub
    Select Case grdLoc.Visible
         Case True
            grdLoc.Visible = False
         Case False
            grdLoc.Rows = 0
            grdLoc.Visible = True
            ActiveReadServer "Select * from Locations order by Location_No"
            While Not rs.EOF
                grdLoc.Rows = grdLoc.Rows + 1
                grdLoc.TextMatrix(grdLoc.Rows - 1, 0) = rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            If grdLoc.Rows > 0 Then
                grdLoc.Height = 650 * grdLoc.Rows - 1
                If grdLoc.Rows > 0 Then grdLoc.Row = 0
                grdLoc.SetFocus
            Else
                lblCashupInfo.Caption = "No Cash'up Available"
                grdLoc.Visible = False
            End If
    End Select
    Me.SetFocus
End Sub

Private Sub Form_Activate()
    lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
    If frmStockTake.Tag = "Not Now" Then
        frmStockTake.Tag = ""
        DoEvents
        Exit Sub
    End If
    If Me.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.785
            Me.Controls(i).Left = Me.Controls(i).Left * 0.782
            Me.Controls(i).Height = Me.Controls(i).Height * 0.78
            Me.Controls(i).top = Me.Controls(i).top * 0.79
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
            Me.Controls(i).FontTextCB.Size = Int(Me.Controls(i).FontTextCB.Size * 0.85)
        Next i
        On Error GoTo 0
        newBack.Width = Me.Width
        newBack.Height = Me.Height
        grdRem.FontSize = 10
        grdMain.FontSize = 10
    End If
    grdMain.ColWidth(0) = grdMain.Width * 0.7
    grdMain.ColWidth(1) = grdMain.Width * 0.3
    grdRem.ColWidth(0) = grdRem.Width * 0.7
    grdRem.ColWidth(1) = grdRem.Width * 0.3
    
    QAnswer.Caption = ""
    lblDetail.Caption = "Select a Location to Request From"
    For i = 0 To cmdDeptStrip.Count - 1
        cmdDeptStrip(i).Visible = False
    Next i
    cmdUser.Caption = "<Request From>"
    grdMain.Rows = 1
    For i = 0 To cmdPlu.Count - 1
        cmdPlu(i).Visible = False
    Next i
End Sub
Private Sub cmdArrow_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    scrolTimer.Enabled = False
    scrolTimer.Interval = 700
End Sub
Private Sub Form_Load()
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    lblDetail.Caption = "Select a Location to Request From"
    For i = 1 To cmdDeptStrip.Count - 1
        cmdDeptStrip(i).Visible = False
    Next i
    Screen.MousePointer = 0
    picSlip.Width = 6485
    grdMain.TextMatrix(0, 0) = "Description"
    grdMain.TextMatrix(0, 1) = "Qty"
    grdRem.TextMatrix(0, 0) = "Description"
    grdRem.TextMatrix(0, 1) = "Qty"
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignRightCenter
    grdRem.ColAlignment(0) = flexAlignLeftCenter
    grdRem.ColAlignment(1) = flexAlignRightCenter
    grdMain.ColHidden(2) = True
    grdMain.ColHidden(3) = True
    grdRem.ColHidden(2) = True
    grdRem.ColHidden(3) = True
End Sub
Private Sub grdLoc_Click()
    cmdUser.Caption = grdLoc.TextMatrix(grdLoc.Row, 0)
    lblDetail.Caption = ""
    grdLoc.Visible = False
    For i = 0 To cmdPlu.Count - 1
        cmdPlu(i).Visible = False
    Next i
    lblDepart.Tag = ""
    Load_Departments Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
    ActiveReadServer "Select * from Location_Stock where Location_No=" & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
    If rs.RecordCount > 0 Then
        lblDepart.Caption = cmdUser.Caption & " > " & rs.Fields("Product_Count") & " Requested of " & rs.Fields("Product_List")
    End If
    rs.Close
End Sub

Private Sub Load_Departments(Location_No)
    For i = 0 To cmdDeptStrip.Count - 1
        cmdDeptStrip(i).Value = 0
    Next i
    grdDept.Rows = 0
    cmdDeptStrip(0).Caption = ""
    cmdDeptStrip(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Departments_Stock where Location_No = " & Location_No & " ORDER BY Short_Name"
    i = -1
    b = 0
    If rs.RecordCount > 1 Then
        Parent = rs.Fields("Dept_Parent")
    End If
    While Not rs.EOF
        i = i + 1
        grdDept.Rows = grdDept.Rows + 1
        If i < 15 And Not rs.EOF Then
            cmdDeptStrip(i).Caption = UCase(Replace(rs.Fields("Short_Name"), "&", "&&"))
            cmdDeptStrip(i).Tag = rs.Fields("Department_no")
            cmdDeptStrip(i).TextDescrCB.Text = rs.Fields("Product_Count") & " Requested of " & rs.Fields("Product_List")
            cmdDeptStrip(i).TextDescrCB.ColorNormal = vbYellow
            cmdDeptStrip(i).TextDescrCB.OffsetY = -6
            If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
            grdDept.Row = grdDept.Rows - 1
            grdDept.TextMatrix(grdDept.Rows - 1, 0) = UCase(Replace(rs.Fields("Short_Name"), "&", "&&"))
            grdDept.TextMatrix(grdDept.Rows - 1, 1) = rs.Fields("Department_No")
            grdDept.TextMatrix(grdDept.Rows - 1, 2) = rs.Fields("Dept_Parent")
            grdDept.TextMatrix(grdDept.Rows - 1, 3) = rs.Fields("Product_Count") & " Requested of " & rs.Fields("Product_List")
        Else
            If b = 0 Then
                grdDept.TextMatrix(grdDept.Rows - 1, 1) = "ßArrow"
                grdDept.Rows = grdDept.Rows + 1
                If i = 15 Then
                    cmdDeptStrip(15).Caption = ""
                    cmdDeptStrip(15).TextDescrCB.Text = ""
                    cmdDeptStrip(15).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdDeptStrip(15).Visible = False Then cmdDeptStrip(15).Visible = True
                End If
            End If
            b = b + 1
            grdDept.TextMatrix(grdDept.Rows - 1, 0) = UCase(Replace(rs.Fields("Short_Name"), "&", "&&"))
            grdDept.TextMatrix(grdDept.Rows - 1, 1) = rs.Fields("Department_No")
            grdDept.TextMatrix(grdDept.Rows - 1, 2) = rs.Fields("Dept_Parent")
            grdDept.TextMatrix(grdDept.Rows - 1, 3) = rs.Fields("Product_Count") & " Requested of " & rs.Fields("Product_List")
            If b = 14 Then b = 0
        End If
        rs.MoveNext
    Wend
    rs.Close
    For b = i + 1 To cmdDeptStrip.Count - 1
       cmdDeptStrip(b).Caption = "1"
       cmdDeptStrip(b).Tag = ""
       cmdDeptStrip(b).Visible = False
    Next b
End Sub
Private Sub Timer1_Timer()
    lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
End Sub

Private Sub Timer2_Timer()
    Select Case cmdErr.BackColor
        Case &HF2&      'White
            cmdErr.BackColor = &HFFFF&
        Case &HFFFF&    'Yellow
            cmdErr.BackColor = &HF2&
    End Select
End Sub
