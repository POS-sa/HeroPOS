VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmRecipe 
   Appearance      =   0  'Flat
   BackColor       =   &H00FAF2F1&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6855
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid grdMess 
      Height          =   3705
      Left            =   5910
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   2865
      _cx             =   5054
      _cy             =   6535
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   975
      Visible         =   0   'False
      Width           =   5850
      _Version        =   524298
      _ExtentX        =   10319
      _ExtentY        =   1720
      _StockProps     =   66
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
      Surface         =   1
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRecipe.frx":0000
      textLT          =   "frmRecipe.frx":0018
      textCT          =   "frmRecipe.frx":0030
      textRT          =   "frmRecipe.frx":0048
      textLM          =   "frmRecipe.frx":0060
      textRM          =   "frmRecipe.frx":0078
      textLB          =   "frmRecipe.frx":0090
      textCB          =   "frmRecipe.frx":00A8
      textRB          =   "frmRecipe.frx":00C0
      colorBack       =   "frmRecipe.frx":00D8
      colorIntern     =   "frmRecipe.frx":0102
      colorMO         =   "frmRecipe.frx":012C
      colorFocus      =   "frmRecipe.frx":0156
      colorDisabled   =   "frmRecipe.frx":0180
      colorPressed    =   "frmRecipe.frx":01AA
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   975
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   1950
      Visible         =   0   'False
      Width           =   5850
      _Version        =   524298
      _ExtentX        =   10319
      _ExtentY        =   1720
      _StockProps     =   66
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
      Surface         =   1
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRecipe.frx":01D4
      textLT          =   "frmRecipe.frx":01EC
      textCT          =   "frmRecipe.frx":0204
      textRT          =   "frmRecipe.frx":021C
      textLM          =   "frmRecipe.frx":0234
      textRM          =   "frmRecipe.frx":024C
      textLB          =   "frmRecipe.frx":0264
      textCB          =   "frmRecipe.frx":027C
      textRB          =   "frmRecipe.frx":0294
      colorBack       =   "frmRecipe.frx":02AC
      colorIntern     =   "frmRecipe.frx":02D6
      colorMO         =   "frmRecipe.frx":0300
      colorFocus      =   "frmRecipe.frx":032A
      colorDisabled   =   "frmRecipe.frx":0354
      colorPressed    =   "frmRecipe.frx":037E
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   975
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   2925
      Visible         =   0   'False
      Width           =   5850
      _Version        =   524298
      _ExtentX        =   10319
      _ExtentY        =   1720
      _StockProps     =   66
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
      Surface         =   1
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRecipe.frx":03A8
      textLT          =   "frmRecipe.frx":03C0
      textCT          =   "frmRecipe.frx":03D8
      textRT          =   "frmRecipe.frx":03F0
      textLM          =   "frmRecipe.frx":0408
      textRM          =   "frmRecipe.frx":0420
      textLB          =   "frmRecipe.frx":0438
      textCB          =   "frmRecipe.frx":0450
      textRB          =   "frmRecipe.frx":0468
      colorBack       =   "frmRecipe.frx":0480
      colorIntern     =   "frmRecipe.frx":04AA
      colorMO         =   "frmRecipe.frx":04D4
      colorFocus      =   "frmRecipe.frx":04FE
      colorDisabled   =   "frmRecipe.frx":0528
      colorPressed    =   "frmRecipe.frx":0552
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   975
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   3900
      Visible         =   0   'False
      Width           =   5850
      _Version        =   524298
      _ExtentX        =   10319
      _ExtentY        =   1720
      _StockProps     =   66
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
      Surface         =   1
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRecipe.frx":057C
      textLT          =   "frmRecipe.frx":0594
      textCT          =   "frmRecipe.frx":05AC
      textRT          =   "frmRecipe.frx":05C4
      textLM          =   "frmRecipe.frx":05DC
      textRM          =   "frmRecipe.frx":05F4
      textLB          =   "frmRecipe.frx":060C
      textCB          =   "frmRecipe.frx":0624
      textRB          =   "frmRecipe.frx":063C
      colorBack       =   "frmRecipe.frx":0654
      colorIntern     =   "frmRecipe.frx":067E
      colorMO         =   "frmRecipe.frx":06A8
      colorFocus      =   "frmRecipe.frx":06D2
      colorDisabled   =   "frmRecipe.frx":06FC
      colorPressed    =   "frmRecipe.frx":0726
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   975
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   4875
      Visible         =   0   'False
      Width           =   5850
      _Version        =   524298
      _ExtentX        =   10319
      _ExtentY        =   1720
      _StockProps     =   66
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
      Surface         =   1
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRecipe.frx":0750
      textLT          =   "frmRecipe.frx":0768
      textCT          =   "frmRecipe.frx":0780
      textRT          =   "frmRecipe.frx":0798
      textLM          =   "frmRecipe.frx":07B0
      textRM          =   "frmRecipe.frx":07C8
      textLB          =   "frmRecipe.frx":07E0
      textCB          =   "frmRecipe.frx":07F8
      textRB          =   "frmRecipe.frx":0810
      colorBack       =   "frmRecipe.frx":0828
      colorIntern     =   "frmRecipe.frx":0852
      colorMO         =   "frmRecipe.frx":087C
      colorFocus      =   "frmRecipe.frx":08A6
      colorDisabled   =   "frmRecipe.frx":08D0
      colorPressed    =   "frmRecipe.frx":08FA
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   945
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Top             =   5850
      Visible         =   0   'False
      Width           =   5850
      _Version        =   524298
      _ExtentX        =   10319
      _ExtentY        =   1667
      _StockProps     =   66
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
      Surface         =   1
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRecipe.frx":0924
      textLT          =   "frmRecipe.frx":093C
      textCT          =   "frmRecipe.frx":0954
      textRT          =   "frmRecipe.frx":096C
      textLM          =   "frmRecipe.frx":0984
      textRM          =   "frmRecipe.frx":099C
      textLB          =   "frmRecipe.frx":09B4
      textCB          =   "frmRecipe.frx":09CC
      textRB          =   "frmRecipe.frx":09E4
      colorBack       =   "frmRecipe.frx":09FC
      colorIntern     =   "frmRecipe.frx":0A26
      colorMO         =   "frmRecipe.frx":0A50
      colorFocus      =   "frmRecipe.frx":0A7A
      colorDisabled   =   "frmRecipe.frx":0AA4
      colorPressed    =   "frmRecipe.frx":0ACE
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCook 
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5850
      _Version        =   524298
      _ExtentX        =   10319
      _ExtentY        =   1720
      _StockProps     =   66
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
      Surface         =   1
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmRecipe.frx":0AF8
      textLT          =   "frmRecipe.frx":0B10
      textCT          =   "frmRecipe.frx":0B28
      textRT          =   "frmRecipe.frx":0B40
      textLM          =   "frmRecipe.frx":0B58
      textRM          =   "frmRecipe.frx":0B70
      textLB          =   "frmRecipe.frx":0B88
      textCB          =   "frmRecipe.frx":0BA0
      textRB          =   "frmRecipe.frx":0BB8
      colorBack       =   "frmRecipe.frx":0BD0
      colorIntern     =   "frmRecipe.frx":0BFA
      colorMO         =   "frmRecipe.frx":0C24
      colorFocus      =   "frmRecipe.frx":0C4E
      colorDisabled   =   "frmRecipe.frx":0C78
      colorPressed    =   "frmRecipe.frx":0CA2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   420
         Top             =   300
      End
   End
   Begin MSForms.Label lblClick 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   5415
      ForeColor       =   12632256
      BackColor       =   16446193
      VariousPropertyBits=   8388627
      Caption         =   "Click to Exit"
      Size            =   "9551;2355"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   555
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCook_Click(Index As Integer)
Screen.MousePointer = 1
    Timer1.Enabled = True
    If cmdCook(Index).Tag = "Exit" Then
        Me.Hide
        Screen.MousePointer = 1
        Exit Sub
    End If
    Screen.MousePointer = 1
    If frmRecipe.Tag = "Sale" Then
        grdMess.Rows = grdMess.Rows + 1
        grdMess.Row = grdMess.Rows - 1
        grdMess.TextMatrix(grdMess.Row, 0) = cmdCook(Index).Tag
        grdMess.TextMatrix(grdMess.Row, 1) = Replace(cmdCook(Index).Caption, "&&", "&")
    End If
FirstLine:
    If Val(cmdCook(Index).Tag) <> 0 Then
        PressClick = False
        ActiveReadServer "Select * from recipes where Product_Code= '" & cmdCook(Index).Tag & "' and Line_Type<>4  order by Line_No"
        If rs.RecordCount > 0 Then
            For i = 0 To cmdCook.Count - 1
                cmdCook(i).Visible = False
                cmdCook(i).Width = frmRecipe.Width
            Next i
            i = -1
            While Not rs.EOF
                i = i + 1
                If i > 6 Then
                    For b = 0 To 6
                        cmdCook(b).Width = frmRecipe.Width / 2
                    Next b
                    If i > cmdCook.UBound Then Load cmdCook(i)
                    cmdCook(i).Visible = True
                    cmdCook(i).Left = frmRecipe.Width / 2
                    cmdCook(i).Width = frmRecipe.Width / 2
                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                End If
                cmdCook(i).Tag = rs.Fields("Line_Code")
                If rs.Fields("Line_Type") = 1 And rs.RecordCount = 1 Then
                    PressClick = True
                Else
                    cmdCook(i).Visible = True
                    cmdCook(i).Caption = UCase(Replace(rs.Fields("Description"), "&", "&&"))
                End If
                rs.MoveNext
            Wend
            rs.Close
            Screen.MousePointer = 1
            If PressClick = True Then
                PressClick = False
                GoTo FirstLine
            End If
         
            Exit Sub
        Else
            rs.Close
        End If
    End If
    Select Case frmRecipe.Tag
        Case "Test"
            If frmProducts.grdMenu.Row < frmProducts.grdMenu.Rows - 1 Then
                frmProducts.grdMenu.Row = frmProducts.grdMenu.Row + 1
                For i = 0 To cmdCook.Count - 1
                            cmdCook(i).Visible = False
                            cmdCook(i).Width = frmRecipe.Width
                        Next i
                i = -1
                If frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 1) <> 0 Then
                    ActiveReadServer "Select * from Recipes where Product_Code= '" & frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 1) & "' order by Line_No"
                    While Not rs.EOF
                        i = i + 1
                        If i > 6 Then
                            For b = 0 To 6
                                cmdCook(b).Width = frmRecipe.Width / 2
                            Next b
                            If i > cmdCook.UBound Then Load cmdCook(i)
                            cmdCook(i).Visible = True
                            cmdCook(i).Left = frmRecipe.Width / 2
                            cmdCook(i).Width = frmRecipe.Width / 2
                            cmdCook(i).top = (i - 7) * cmdCook(i).Height
                        End If
                        cmdCook(i).Visible = True
                        If rs.Fields("Line_Code") = 0 Then
                            Description = Replace(rs.Fields("Description"), "&", "&&")
                        Else
                            Description = Trim(Mid(Replace(rs.Fields("Description"), "&", "&&"), 1, InStrRev(Replace(rs.Fields("Description"), "&", "&&"), ",") - 1))
                            If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                            If rs.Fields("Line_Type") = 3 Then
                                ActiveReadServer1 "Select Short_Description from Products where Product_Code='" & rs.Fields("Line_Code") & "'"
                                If rs1.RecordCount > 0 Then Description = rs1.Fields("Short_Description") & ""
                                rs1.Close
                            End If
                        End If
                        cmdCook(i).Caption = UCase(Description)
                        cmdCook(i).Tag = rs.Fields("Line_Code")
                        rs.MoveNext
                    Wend
                    If rs.RecordCount = 0 Then
                        rs.Close
                        
                        
                        
                        ActiveReadServer "Select Line_Type from Recipes where Line_Code = '" & cmdCook(Index).Tag & "' and Product_Code = '" & frmProducts.txtProductCode & "'"
                        If rs.RecordCount > 0 Then
                            If rs.Fields("Line_Type") = 6 Or rs.Fields("Line_Type") = 7 Then
                                cmdCook(Index).BackColor = &H44A3D0
                                If InStr(cmdCook(Index).Caption, Chr(215)) = 0 Then
                                    cmdCook(Index).Caption = "1 " & Chr(215) & " " & cmdCook(Index).Caption
                                Else
                                    cmdCook(Index).Caption = Val(Mid(cmdCook(Index).Caption, 1, InStr(cmdCook(Index).Caption, Chr(215)) - 1)) + 1 & " " & Chr(215) & " " & Mid(cmdCook(Index).Caption, InStr(cmdCook(Index).Caption, Chr(215)) + 1)
                                End If
                            Else
                                For i = 0 To cmdCook.Count - 1
                                    cmdCook(i).Visible = False
                                    cmdCook(i).Width = frmRecipe.Width
                                Next i
                                Me.Hide
                              
                                Exit Sub
                            End If
                        Else
                            For i = 0 To cmdCook.Count - 1
                                cmdCook(i).Visible = False
                                cmdCook(i).Width = frmRecipe.Width
                            Next i
                            Me.Hide
                         
                            Exit Sub
                        End If
                    End If
                    rs.Close
                Else
top:
                    i = i + 1
                    If i > 6 Then
                        For b = 0 To 6
                            cmdCook(b).Width = frmRecipe.Width / 2
                        Next b
                        If i > cmdCook.UBound Then Load cmdCook(i)
                        cmdCook(i).Visible = True
                        cmdCook(i).Left = frmRecipe.Width / 2
                        cmdCook(i).Width = frmRecipe.Width / 2
                        cmdCook(i).top = (i - 7) * cmdCook(i).Height
                    End If
                    cmdCook(i).Visible = True
                    Description = frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 3)
                    cmdCook(i).Caption = UCase(Description)
                    cmdCook(i).Tag = frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 1)
                    If frmProducts.grdMenu.Row < frmProducts.grdMenu.Rows - 1 Then
                        If frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row + 1, 1) = 0 Then
                            frmProducts.grdMenu.Row = frmProducts.grdMenu.Row + 1
                            GoTo top
                        End If
                    End If
                End If
            Else
                
               
                ActiveReadServer1 "Select Line_Type from Recipes where Line_Code = '" & cmdCook(Index).Tag & "' and Product_Code = '" & frmProducts.txtProductCode & "'"
                If rs1.RecordCount > 0 Then
                    If rs1.Fields("Line_Type") = 6 Or rs1.Fields("Line_Type") = 7 Then
                        cmdCook(Index).BackColor = &H44A3D0
                        If InStr(cmdCook(Index).Caption, Chr(215)) = 0 Then
                            cmdCook(Index).Caption = "1 " & Chr(215) & " " & cmdCook(Index).Caption
                        Else
                            cmdCook(Index).Caption = Val(Mid(cmdCook(Index).Caption, 1, InStr(cmdCook(Index).Caption, Chr(215)) - 1)) + 1 & " " & Chr(215) & " " & Mid(cmdCook(Index).Caption, InStr(cmdCook(Index).Caption, Chr(215)) + 1)
                        End If
                    Else
                        Me.Hide
                       
                        Exit Sub
                    End If
                Else
                    Me.Hide
                    
                    Exit Sub
                End If
            End If
        Case "Sale"
            Select Case Panel_no
                Case 0
                    If frmSales.grdMenu.Row < frmSales.grdMenu.Rows - 1 Then
                        frmSales.grdMenu.Row = frmSales.grdMenu.Row + 1
                        For i = 0 To cmdCook.Count - 1
                            cmdCook(i).Visible = False
                            cmdCook(i).Width = frmRecipe.Width
                        Next i
                        i = -1
                        If frmSales.grdMenu.ValueMatrix(frmSales.grdMenu.Row, 1) <> 0 Then
                            ActiveReadServer "Select * from recipes where Product_Code= '" & frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 1) & "' order by Line_No"
                            frmSales.lblKeyRegister.TextAlign = fmTextAlignLeft
                            frmSales.lblKeyRegister = Mid(frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 3), 1, InStrRev(frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 3), ",") - 1)
                            While Not rs.EOF
                                i = i + 1
                                If i > 6 Then
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 2
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    cmdCook(i).Width = frmRecipe.Width / 2
                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                End If
                                cmdCook(i).Visible = True
                                If rs.Fields("Line_Code") = 0 Then
                                    Description = Replace(rs.Fields("Description"), "&", "&&")
                                Else
                                    Description = Trim(Mid(Replace(rs.Fields("Description"), "&", "&&"), 1, InStrRev(Replace(rs.Fields("Description"), "&", "&&"), ",") - 1))
                                    If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                                End If
                                cmdCook(i).Caption = UCase(Description)
                                cmdCook(i).Tag = rs.Fields("Line_Code")
                                rs.MoveNext
                            Wend
                            If rs.RecordCount = 0 Then
                                ActiveReadServer "Select Line_Type from Recipes where Line_Code = '" & cmdCook(Index).Tag & "' and Product_Code = '" & TillData.ProductCode & "'"
                                If rs.RecordCount > 0 Then
                                    If rs.Fields("Line_Type") = 6 Or rs.Fields("Line_Type") = 7 Then
                                        cmdCook(Index).BackColor = &H44A3D0
                                        If InStr(cmdCook(Index).Caption, Chr(215)) = 0 Then
                                            cmdCook(Index).Caption = "1 " & Chr(215) & " " & cmdCook(Index).Caption
                                        Else
                                            cmdCook(Index).Caption = Val(Mid(cmdCook(Index).Caption, 1, InStr(cmdCook(Index).Caption, Chr(215)) - 1)) + 1 & " " & Chr(215) & " " & Mid(cmdCook(Index).Caption, InStr(cmdCook(Index).Caption, Chr(215)) + 1)
                                        End If
                                    Else
                                        For i = 0 To cmdCook.Count - 1
                                            cmdCook(i).Visible = False
                                            cmdCook(i).Width = frmRecipe.Width
                                        Next i
                                        Me.Hide
                                       
                                        Exit Sub
                                    End If
                                Else
                                    For i = 0 To cmdCook.Count - 1
                                        cmdCook(i).Visible = False
                                        cmdCook(i).Width = frmRecipe.Width
                                    Next i
                                    Me.Hide
                                  
                                    Exit Sub
                                End If
                            End If
                            rs.Close
                        Else
top1:
                            i = i + 1
                            If i > 6 Then
                                For b = 0 To 6
                                    cmdCook(b).Width = frmRecipe.Width / 2
                                Next b
                                If i > cmdCook.UBound Then Load cmdCook(i)
                                cmdCook(i).Visible = True
                                cmdCook(i).Left = frmRecipe.Width / 2
                                cmdCook(i).Width = frmRecipe.Width / 2
                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                            End If
                            cmdCook(i).Visible = True
                            Description = frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 3)
                            cmdCook(i).Caption = UCase(Description)
                            cmdCook(i).Tag = frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 1)
                            If frmSales.grdMenu.Row < frmSales.grdMenu.Rows - 1 Then
                                If Val(frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row + 1, 1)) = 0 Then
                                    frmSales.grdMenu.Row = frmSales.grdMenu.Row + 1
                                    GoTo top1
                                End If
                            End If
                        End If
                    Else
                        ActiveReadServer "Select Line_Type from Recipes where Line_Code = '" & cmdCook(Index).Tag & "' and Product_Code = '" & TillData.ProductCode & "'"
                        If rs.RecordCount > 0 Then
                            If rs.Fields("Line_Type") = 6 Or rs.Fields("Line_Type") = 7 Then
                                cmdCook(Index).BackColor = &H44A3D0
                                If InStr(cmdCook(Index).Caption, Chr(215)) = 0 Then
                                    cmdCook(Index).Caption = "1 " & Chr(215) & " " & cmdCook(Index).Caption
                                Else
                                    cmdCook(Index).Caption = Val(Mid(cmdCook(Index).Caption, 1, InStr(cmdCook(Index).Caption, Chr(215)) - 1)) + 1 & " " & Chr(215) & " " & Mid(cmdCook(Index).Caption, InStr(cmdCook(Index).Caption, Chr(215)) + 1)
                                End If
                            Else
                                Me.Hide
                              
                                Exit Sub
                            End If
                        Else
                            Me.Hide
                         
                            Exit Sub
                        End If
                    End If
                Case 1
                    If frmSales1.grdMenu.Row < frmSales1.grdMenu.Rows - 1 Then
                        frmSales1.grdMenu.Row = frmSales1.grdMenu.Row + 1
                        For i = 0 To cmdCook.Count - 1
                            cmdCook(i).Visible = False
                            cmdCook(i).Width = frmRecipe.Width
                        Next i
                        i = -1
                        If frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 1) <> 0 Then
                            ActiveReadServer "Select * from recipes where Product_Code= '" & frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 1) & "' order by Line_No"
                            frmSales1.lblKeyRegister.TextAlign = fmTextAlignLeft
                            frmSales1.lblKeyRegister = Mid(frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 3), 1, InStrRev(frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 3), ",") - 1)
                            While Not rs.EOF
                                i = i + 1
                                If i > 6 Then
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 2
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    cmdCook(i).Width = frmRecipe.Width / 2
                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                End If
                                cmdCook(i).Visible = True
                                If rs.Fields("Line_Code") = 0 Then
                                    Description = Replace(rs.Fields("Description"), "&", "&&")
                                Else
                                    Description = Trim(Mid(Replace(rs.Fields("Description"), "&", "&&"), 1, InStrRev(Replace(rs.Fields("Description"), "&", "&&"), ",") - 1))
                                    If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                                    If rs.Fields("Line_Type") = 3 Then
                                        ActiveReadServer1 "Select Short_Description from Products where Product_Code='" & rs.Fields("Line_Code") & "'"
                                        If rs1.RecordCount > 0 Then Description = rs1.Fields("Short_Description") & ""
                                        rs1.Close
                                    End If
                                End If
                                cmdCook(i).Caption = UCase(Description)
                                cmdCook(i).Tag = rs.Fields("Line_Code")
                                rs.MoveNext
                            Wend
                            If rs.RecordCount = 0 Then
                                rs.Close
                                Me.Hide
                             
                                Exit Sub
                            End If
                            rs.Close
                        Else
top2:
                            i = i + 1
                            If i > 6 Then
                                For b = 0 To 6
                                    cmdCook(b).Width = frmRecipe.Width / 2
                                Next b
                                If i > cmdCook.UBound Then Load cmdCook(i)
                                cmdCook(i).Visible = True
                                cmdCook(i).Left = frmRecipe.Width / 2
                                cmdCook(i).Width = frmRecipe.Width / 2
                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                            End If
                            cmdCook(i).Visible = True
                            Description = frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 3)
                            cmdCook(i).Caption = UCase(Description)
                            cmdCook(i).Tag = frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 1)
                            If frmSales1.grdMenu.Row < frmSales1.grdMenu.Rows - 1 Then
                                If frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row + 1, 1) = 0 Then
                                    frmSales1.grdMenu.Row = frmSales1.grdMenu.Row + 1
                                    GoTo top2
                                End If
                            End If
                        End If
                    Else
                        ActiveReadServer "Select Line_Type,Cost from Recipes where Line_Code = '" & cmdCook(Index).Tag & "' and Product_Code = '" & TillData.ProductCode & "'"
                        If rs.RecordCount > 0 Then
                            If rs.Fields("Line_Type") = 6 Or rs.Fields("Line_Type") = 7 Then
                                cmdCook(Index).BackColor = &H44A3D0
                                If InStr(cmdCook(Index).Caption, Chr(215)) = 0 Then
                                    cmdCook(Index).Caption = "1 " & Chr(215) & " " & cmdCook(Index).Caption
                                    If rs.Fields("Line_Type") = 6 Then
                                        TillData.Price = TillData.Price + rs.Fields("Cost")
                                    End If
                                Else
                                    cmdCook(Index).Caption = Val(Mid(cmdCook(Index).Caption, 1, InStr(cmdCook(Index).Caption, Chr(215)) - 1)) + 1 & " " & Chr(215) & " " & Mid(cmdCook(Index).Caption, InStr(cmdCook(Index).Caption, Chr(215)) + 1)
                                End If
                            Else
                                Me.Hide
                          
                                Exit Sub
                            End If
                        Else
                            Me.Hide
                       
                            Exit Sub
                        End If
                    End If
                Case 2
                    If frmBar.grdMenu.Row < frmBar.grdMenu.Rows - 1 Then
                        frmBar.grdMenu.Row = frmBar.grdMenu.Row + 1
                        For i = 0 To cmdCook.Count - 1
                            cmdCook(i).Visible = False
                            cmdCook(i).Width = frmRecipe.Width
                        Next i
                        i = -1
                        If frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 1) <> 0 Then
                            ActiveReadServer "Select * from recipes where Product_Code= '" & frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 1) & "' order by Line_No"
                            frmBar.lblKeyRegister.TextAlign = fmTextAlignLeft
                            frmBar.lblKeyRegister = Mid(frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 3), 1, InStrRev(frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 3), ",") - 1)
                            While Not rs.EOF
                                i = i + 1
                                If i > 6 Then
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 2
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    cmdCook(i).Width = frmRecipe.Width / 2
                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                End If
                                cmdCook(i).Visible = True
                                If rs.Fields("Line_Code") = 0 Then
                                    Description = Replace(rs.Fields("Description"), "&", "&&")
                                Else
                                    Description = Trim(Mid(Replace(rs.Fields("Description"), "&", "&&"), 1, InStrRev(Replace(rs.Fields("Description"), "&", "&&"), ",") - 1))
                                    If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                                End If
                                cmdCook(i).Caption = UCase(Description)
                                cmdCook(i).Tag = rs.Fields("Line_Code")
                                rs.MoveNext
                            Wend
                            If rs.RecordCount = 0 Then
                                rs.Close
                                Me.Hide
                             
                                Exit Sub
                            End If
                            rs.Close
                        Else
top3:
                            i = i + 1
                            If i > 6 Then
                                For b = 0 To 6
                                    cmdCook(b).Width = frmRecipe.Width / 2
                                Next b
                                If i > cmdCook.UBound Then Load cmdCook(i)
                                cmdCook(i).Visible = True
                                cmdCook(i).Left = frmRecipe.Width / 2
                                cmdCook(i).Width = frmRecipe.Width / 2
                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                            End If
                            cmdCook(i).Visible = True
                            Description = frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 3)
                            cmdCook(i).Caption = UCase(Description)
                            cmdCook(i).Tag = frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 1)
                            If frmBar.grdMenu.Row < frmBar.grdMenu.Rows - 1 Then
                                If frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row + 1, 1) = 0 Then
                                    frmBar.grdMenu.Row = frmBar.grdMenu.Row + 1
                                    GoTo top3
                                End If
                            End If
                        End If
                    Else
                        ActiveReadServer "Select Line_Type from Recipes where Line_Code = '" & cmdCook(Index).Tag & "' and Product_Code = '" & TillData.ProductCode & "'"
                        If rs.RecordCount > 0 Then
                            If rs.Fields("Line_Type") = 6 Or rs.Fields("Line_Type") = 7 Then
                                cmdCook(Index).BackColor = &H44A3D0
                                If InStr(cmdCook(Index).Caption, Chr(215)) = 0 Then
                                    cmdCook(Index).Caption = "1 " & Chr(215) & " " & cmdCook(Index).Caption
                                Else
                                    cmdCook(Index).Caption = Val(Mid(cmdCook(Index).Caption, 1, InStr(cmdCook(Index).Caption, Chr(215)) - 1)) + 1 & " " & Chr(215) & " " & Mid(cmdCook(Index).Caption, InStr(cmdCook(Index).Caption, Chr(215)) + 1)
                                End If
                            Else
                                Me.Hide
                              
                                Exit Sub
                            End If
                        Else
                            Me.Hide
                            
                            Exit Sub
                        End If
                    End If
            End Select
    End Select

End Sub
Private Sub Form_Activate()
    Timer1.Enabled = True
    Screen.MousePointer = 1
    If cmdCook.UBound > 6 Then
        For i = 7 To cmdCook.UBound
            Unload cmdCook(i)
        Next i
    End If
    For i = 0 To cmdCook.Count - 1
        cmdCook(i).Visible = False
        cmdCook(i).Width = frmRecipe.Width
        cmdCook(i).BackColor = &HA3D3E9
        cmdCook(1).FontTextCaption.Size = 14
    Next i
    grdMess.Rows = 0
    i = -1
top:
    Select Case frmRecipe.Tag
        Case "Test"
            If frmProducts.grdMenu.Rows > 0 Then
                Select Case frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 2)
                    Case 0, 8
                        i = i + 1
                        Select Case i
                            Case Is > 12
                                For b = 0 To 6
                                    cmdCook(b).Width = frmRecipe.Width / 3
                                Next b
                                For b = 7 To 12
                                    cmdCook(b).Width = frmRecipe.Width / 3
                                    cmdCook(b).Left = cmdCook(0).Width
                                Next b
                                If i > cmdCook.UBound Then Load cmdCook(i)
                                cmdCook(i).Visible = True
                                cmdCook(i).Left = frmRecipe.Width / 2
                                Select Case i
                                    Case Is > 13
                                        cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                        cmdCook(i).top = cmdCook(i - 7).top
                                    Case 13
                                        cmdCook(i).Left = cmdCook(0).Width
                                        cmdCook(i).top = cmdCook(6).top
                                    Case Else
                                        cmdCook(i).Width = frmRecipe.Width / 2
                                        cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                End Select
                                If i > 13 Then
                                    For b = 0 To i
                                        cmdCook(b).FontTextCaption.Size = 11
                                    Next b
                                Else
                                    For b = 0 To i
                                        cmdCook(b).FontTextCaption.Size = 14
                                    Next b
                                End If
                            Case Is > 6
                                For b = 0 To 6
                                    cmdCook(b).Width = frmRecipe.Width / 2
                                Next b
                                If i > cmdCook.UBound Then Load cmdCook(i)
                                cmdCook(i).Visible = True
                                cmdCook(i).Left = frmRecipe.Width / 2
                                cmdCook(i).Width = frmRecipe.Width / 2
                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                        End Select
                        cmdCook(i).Visible = True
                        If frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 2) = 8 Then
                            cmdCook(i).Caption = "EXIT"
                            cmdCook(i).Tag = "Exit"
                            cmdCook(i).BackColor = &H80C0FF
                        Else
                            cmdCook(i).Caption = UCase(frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 3))
                            cmdCook(i).Tag = frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 1)
                        End If
                        If frmProducts.grdMenu.Row <> frmProducts.grdMenu.Rows - 1 Then
                            frmProducts.grdMenu.Row = frmProducts.grdMenu.Row + 1
                            GoTo top
                        End If
                    Case 2, 3, 6, 7
                        i = i + 1
                        Select Case i
                            Case Is > 12
                                For b = 0 To 6
                                    cmdCook(b).Width = frmRecipe.Width / 3
                                Next b
                                For b = 7 To 12
                                    cmdCook(b).Width = frmRecipe.Width / 3
                                    cmdCook(b).Left = cmdCook(0).Width
                                Next b
                                If i > cmdCook.UBound Then Load cmdCook(i)
                                cmdCook(i).Visible = True
                                cmdCook(i).Left = frmRecipe.Width / 2
                                Select Case i
                                    Case Is > 13
                                        cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                        cmdCook(i).top = cmdCook(i - 7).top
                                    Case 13
                                        cmdCook(i).Left = cmdCook(0).Width
                                        cmdCook(i).top = cmdCook(6).top
                                    Case Else
                                        cmdCook(i).Width = frmRecipe.Width / 2
                                        cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                End Select
                                If i > 13 Then
                                    For b = 0 To i
                                        cmdCook(b).FontTextCaption.Size = 11
                                    Next b
                                Else
                                    For b = 0 To i
                                        cmdCook(b).FontTextCaption.Size = 14
                                    Next b
                                End If
                            Case Is > 6
                                For b = 0 To 6
                                    cmdCook(b).Width = frmRecipe.Width / 2
                                Next b
                                If i > cmdCook.UBound Then Load cmdCook(i)
                                cmdCook(i).Visible = True
                                cmdCook(i).Left = frmRecipe.Width / 2
                                cmdCook(i).Width = frmRecipe.Width / 2
                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                        End Select
                        cmdCook(i).Visible = True
                        If frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 2) = 6 Or frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 2) = 7 Or frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 2) = 3 Then
                            ActiveReadServer1 "Select Short_Description from Products where Product_Code='" & frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 1) & "'"
                            If rs1.RecordCount > 0 Then Description = rs1.Fields("Short_Description") & ""
                            rs1.Close
                        Else
                            Description = Trim(Mid(frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 3), 1, InStrRev(frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 3), ",") - 1))
                            If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                        End If
                        cmdCook(i).Caption = UCase(Description)
                        cmdCook(i).Tag = frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 1)
                        If frmProducts.grdMenu.Row <> frmProducts.grdMenu.Rows - 1 Then
                            frmProducts.grdMenu.Row = frmProducts.grdMenu.Row + 1
                            GoTo top
                        End If
                    Case 1
                        If i = -1 Then
                            ActiveReadServer "Select * from recipes where Product_Code= '" & frmProducts.grdMenu.TextMatrix(frmProducts.grdMenu.Row, 1) & "' order by Line_No"
                            While Not rs.EOF
                                i = i + 1
                                Select Case i
                                    Case Is > 12
                                        For b = 0 To 6
                                            cmdCook(b).Width = frmRecipe.Width / 3
                                        Next b
                                        For b = 7 To 12
                                            cmdCook(b).Width = frmRecipe.Width / 3
                                            cmdCook(b).Left = cmdCook(0).Width
                                        Next b
                                        If i > cmdCook.UBound Then Load cmdCook(i)
                                        cmdCook(i).Visible = True
                                        cmdCook(i).Left = frmRecipe.Width / 2
                                        Select Case i
                                            Case Is > 13
                                                cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                                cmdCook(i).top = cmdCook(i - 7).top
                                            Case 13
                                                cmdCook(i).Left = cmdCook(0).Width
                                                cmdCook(i).top = cmdCook(6).top
                                            Case Else
                                                cmdCook(i).Width = frmRecipe.Width / 2
                                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                        End Select
                                        If i > 13 Then
                                            For b = 0 To i
                                                cmdCook(b).FontTextCaption.Size = 11
                                            Next b
                                        Else
                                            For b = 0 To i
                                                cmdCook(b).FontTextCaption.Size = 14
                                            Next b
                                        End If
                                    Case Is > 6
                                        For b = 0 To 6
                                            cmdCook(b).Width = frmRecipe.Width / 2
                                        Next b
                                        If i > cmdCook.UBound Then Load cmdCook(i)
                                        cmdCook(i).Visible = True
                                        cmdCook(i).Left = frmRecipe.Width / 2
                                        cmdCook(i).Width = frmRecipe.Width / 2
                                        cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                End Select
                                cmdCook(i).Visible = True
                                If rs.Fields("Line_Code") = 0 Then
                                    cmdCook(i).Caption = UCase(rs.Fields("Description"))
                                Else
                                
                                Description = Trim(Mid(rs.Fields("Description"), 1, InStrRev(rs.Fields("Description"), ",") - 1))
                                    If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                                    cmdCook(i).Caption = UCase(Description)
                                End If
                                cmdCook(i).Tag = rs.Fields("Line_Code")
                                rs.MoveNext
                            Wend
                            rs.Close
                        End If
                End Select
            End If
        Case "Sale"
            Select Case Panel_no
                Case 0, 8
                    If frmSales.grdMenu.Rows > 0 Then
                        Select Case frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 2)
                            Case 0
                                i = i + 1
                                Select Case i
                                    Case Is > 12
                                        For b = 0 To 6
                                            cmdCook(b).Width = frmRecipe.Width / 3
                                        Next b
                                        For b = 7 To 12
                                            cmdCook(b).Width = frmRecipe.Width / 3
                                            cmdCook(b).Left = cmdCook(0).Width
                                        Next b
                                        If i > cmdCook.UBound Then Load cmdCook(i)
                                        cmdCook(i).Visible = True
                                        cmdCook(i).Left = frmRecipe.Width / 2
                                        Select Case i
                                            Case Is > 13
                                                cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                                cmdCook(i).top = cmdCook(i - 7).top
                                            Case 13
                                                cmdCook(i).Left = cmdCook(0).Width
                                                cmdCook(i).top = cmdCook(6).top
                                            Case Else
                                                cmdCook(i).Width = frmRecipe.Width / 2
                                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                        End Select
                                        If i > 13 Then
                                            For b = 0 To i
                                                cmdCook(b).FontTextCaption.Size = 11
                                            Next b
                                        Else
                                            For b = 0 To i
                                                cmdCook(b).FontTextCaption.Size = 14
                                            Next b
                                        End If
                                    Case Is > 6
                                        For b = 0 To 6
                                            cmdCook(b).Width = frmRecipe.Width / 2
                                        Next b
                                        If i > cmdCook.UBound Then Load cmdCook(i)
                                        cmdCook(i).Visible = True
                                        cmdCook(i).Left = frmRecipe.Width / 2
                                        cmdCook(i).Width = frmRecipe.Width / 2
                                        cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                End Select
                                cmdCook(i).Visible = True
                                If frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 2) = 8 Then
                                    cmdCook(i).Caption = "EXIT"
                                    cmdCook(i).Tag = "Exit"
                                    cmdCook(i).BackColor = &H80C0FF
                                Else
                                    cmdCook(i).Caption = UCase(frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 3))
                                    cmdCook(i).Tag = frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 1)
                                End If
                                If frmSales.grdMenu.Row <> frmSales.grdMenu.Rows - 1 Then
                                    frmSales.grdMenu.Row = frmSales.grdMenu.Row + 1
                                    GoTo top
                                End If
                            Case 2, 3, 6, 7
                                i = i + 1
                                Select Case i
                                    Case Is > 12
                                        For b = 0 To 6
                                            cmdCook(b).Width = frmRecipe.Width / 3
                                        Next b
                                        For b = 7 To 12
                                            cmdCook(b).Width = frmRecipe.Width / 3
                                            cmdCook(b).Left = cmdCook(0).Width
                                        Next b
                                        If i > cmdCook.UBound Then Load cmdCook(i)
                                        cmdCook(i).Visible = True
                                        cmdCook(i).Left = frmRecipe.Width / 2
                                        Select Case i
                                            Case Is > 13
                                                cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                                cmdCook(i).top = cmdCook(i - 7).top
                                            Case 13
                                                cmdCook(i).Left = cmdCook(0).Width
                                                cmdCook(i).top = cmdCook(6).top
                                            Case Else
                                                cmdCook(i).Width = frmRecipe.Width / 2
                                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                        End Select
                                        If i > 13 Then
                                            For b = 0 To i
                                                cmdCook(b).FontTextCaption.Size = 11
                                            Next b
                                        Else
                                            For b = 0 To i
                                                cmdCook(b).FontTextCaption.Size = 14
                                            Next b
                                        End If
                                    Case Is > 6
                                        For b = 0 To 6
                                            cmdCook(b).Width = frmRecipe.Width / 2
                                        Next b
                                        If i > cmdCook.UBound Then Load cmdCook(i)
                                        cmdCook(i).Visible = True
                                        cmdCook(i).Left = frmRecipe.Width / 2
                                        cmdCook(i).Width = frmRecipe.Width / 2
                                        cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                End Select
                                cmdCook(i).Visible = True
                                If frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 2) = 6 Or frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 2) = 7 Or frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 2) = 3 Then
                                    ActiveReadServer1 "Select Short_Description from Products where Product_Code='" & frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 1) & "'"
                                    If rs1.RecordCount > 0 Then Description = rs1.Fields("Short_Description") & ""
                                    rs1.Close
                                Else
                                    Description = Trim(Mid(frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 3), 1, InStrRev(frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 3), ",") - 1))
                                    If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                                End If
                                If frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 2) = 6 Or frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 2) = 7 Then
                                    cmdCook(i).Caption = UCase(Description & Chr(160))
                                Else
                                    cmdCook(i).Caption = UCase(Description)
                                End If
                                cmdCook(i).Tag = frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 1)
                                If frmSales.grdMenu.Row <> frmSales.grdMenu.Rows - 1 Then
                                    frmSales.grdMenu.Row = frmSales.grdMenu.Row + 1
                                    GoTo top
                                End If
                            Case 1
                                If i = -1 Then
                                    ActiveReadServer "Select * from recipes where Product_Code= '" & frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 1) & "' order by Line_No"
                                    frmSales.lblKeyRegister.TextAlign = fmTextAlignLeft
                                    frmSales.lblKeyRegister = Mid(frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 3), 1, InStrRev(frmSales.grdMenu.TextMatrix(frmSales.grdMenu.Row, 3), ",") - 1)
                                    While Not rs.EOF
                                        i = i + 1
                                        Select Case i
                                            Case Is > 12
                                                For b = 0 To 6
                                                    cmdCook(b).Width = frmRecipe.Width / 3
                                                Next b
                                                For b = 7 To 12
                                                    cmdCook(b).Width = frmRecipe.Width / 3
                                                    cmdCook(b).Left = cmdCook(0).Width
                                                Next b
                                                If i > cmdCook.UBound Then Load cmdCook(i)
                                                cmdCook(i).Visible = True
                                                cmdCook(i).Left = frmRecipe.Width / 2
                                                Select Case i
                                                    Case Is > 13
                                                        cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                                        cmdCook(i).top = cmdCook(i - 7).top
                                                    Case 13
                                                        cmdCook(i).Left = cmdCook(0).Width
                                                        cmdCook(i).top = cmdCook(6).top
                                                    Case Else
                                                        cmdCook(i).Width = frmRecipe.Width / 2
                                                        cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                                End Select
                                                If i > 13 Then
                                                    For b = 0 To i
                                                        cmdCook(b).FontTextCaption.Size = 11
                                                    Next b
                                                Else
                                                    For b = 0 To i
                                                        cmdCook(b).FontTextCaption.Size = 14
                                                    Next b
                                                End If
                                            Case Is > 6
                                                For b = 0 To 6
                                                    cmdCook(b).Width = frmRecipe.Width / 2
                                                Next b
                                                If i > cmdCook.UBound Then Load cmdCook(i)
                                                cmdCook(i).Visible = True
                                                cmdCook(i).Left = frmRecipe.Width / 2
                                                cmdCook(i).Width = frmRecipe.Width / 2
                                                cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                        End Select
                                        cmdCook(i).Visible = True
                                        If InStr(rs.Fields("Description"), ",") <> 0 Then
                                            Description = Trim(Mid(rs.Fields("Description"), 1, InStrRev(rs.Fields("Description"), ",") - 1))
                                        Else
                                            Description = rs.Fields("Description")
                                        End If
                                        If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                                        cmdCook(i).Caption = UCase(Description)
                                        cmdCook(i).Tag = rs.Fields("Line_Code")
                                        rs.MoveNext
                                    Wend
                                    rs.Close
                                End If
                        End Select
                    End If
            Case 1
                If frmSales1.grdMenu.Rows > 0 Then
                    Select Case frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 2)
                        Case 0, 8
                            i = i + 1
                            Select Case i
                                Case Is > 12
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 3
                                    Next b
                                    For b = 7 To 12
                                        cmdCook(b).Width = frmRecipe.Width / 3
                                        cmdCook(b).Left = cmdCook(0).Width
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    Select Case i
                                        Case Is > 13
                                            cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                            cmdCook(i).top = cmdCook(i - 7).top
                                        Case 13
                                            cmdCook(i).Left = cmdCook(0).Width
                                            cmdCook(i).top = cmdCook(6).top
                                        Case Else
                                            cmdCook(i).Width = frmRecipe.Width / 2
                                            cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                    End Select
                                    If i > 13 Then
                                        For b = 0 To i
                                            cmdCook(b).FontTextCaption.Size = 11
                                        Next b
                                    Else
                                        For b = 0 To i
                                            cmdCook(b).FontTextCaption.Size = 14
                                        Next b
                                    End If
                                Case Is > 6
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 2
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    cmdCook(i).Width = frmRecipe.Width / 2
                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                            End Select
                            cmdCook(i).Visible = True
                            If frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 2) = 8 Then
                                cmdCook(i).Caption = "EXIT"
                                cmdCook(i).Tag = "Exit"
                                cmdCook(i).BackColor = &H80C0FF
                            Else
                                cmdCook(i).Caption = UCase(frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 3))
                                cmdCook(i).Tag = frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 1)
                            End If
                            If frmSales1.grdMenu.Row <> frmSales1.grdMenu.Rows - 1 Then
                                frmSales1.grdMenu.Row = frmSales1.grdMenu.Row + 1
                                GoTo top
                            End If
                        Case 2, 3, 6, 7
                            i = i + 1
                            Select Case i
                                Case Is > 12
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 3
                                    Next b
                                    For b = 7 To 12
                                        cmdCook(b).Width = frmRecipe.Width / 3
                                        cmdCook(b).Left = cmdCook(0).Width
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    Select Case i
                                        Case Is > 13
                                            cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                            cmdCook(i).top = cmdCook(i - 7).top
                                        Case 13
                                            cmdCook(i).Left = cmdCook(0).Width
                                            cmdCook(i).top = cmdCook(6).top
                                        Case Else
                                            cmdCook(i).Width = frmRecipe.Width / 2
                                            cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                    End Select
                                    If i > 13 Then
                                        For b = 0 To i
                                            cmdCook(b).FontTextCaption.Size = 11
                                        Next b
                                    Else
                                        For b = 0 To i
                                            cmdCook(b).FontTextCaption.Size = 14
                                        Next b
                                    End If
                                Case Is > 6
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 2
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    cmdCook(i).Width = frmRecipe.Width / 2
                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                            End Select
                            cmdCook(i).Visible = True
                            If frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 2) = 6 Or frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 2) = 7 Or frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 2) = 3 Then
                                ActiveReadServer1 "Select Short_Description from Products where Product_Code='" & frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 1) & "'"
                                If rs1.RecordCount > 0 Then Description = rs1.Fields("Short_Description") & ""
                                rs1.Close
                            Else
                                Description = Trim(Mid(frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 3), 1, InStrRev(frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 3), ",") - 1))
                                If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                            End If
                            If frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 2) = 6 Or frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 2) = 7 Then
                                cmdCook(i).Caption = UCase(Description & Chr(160))
                            Else
                                cmdCook(i).Caption = UCase(Description)
                            End If
                            cmdCook(i).Tag = frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 1)
                            If frmSales1.grdMenu.Row <> frmSales1.grdMenu.Rows - 1 Then
                                frmSales1.grdMenu.Row = frmSales1.grdMenu.Row + 1
                                GoTo top
                            End If
                        Case 1
                            If i = -1 Then
                                ActiveReadServer "Select * from recipes where Product_Code= '" & frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 1) & "' order by Line_No"
                                frmSales1.lblKeyRegister.TextAlign = fmTextAlignLeft
                                frmSales1.lblKeyRegister = Mid(frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 3), 1, InStrRev(frmSales1.grdMenu.TextMatrix(frmSales1.grdMenu.Row, 3), ",") - 1)
                                While Not rs.EOF
                                    i = i + 1
                                    Select Case i
                                        Case Is > 12
                                            For b = 0 To 6
                                                cmdCook(b).Width = frmRecipe.Width / 3
                                            Next b
                                            For b = 7 To 12
                                                cmdCook(b).Width = frmRecipe.Width / 3
                                                cmdCook(b).Left = cmdCook(0).Width
                                            Next b
                                            If i > cmdCook.UBound Then Load cmdCook(i)
                                            cmdCook(i).Visible = True
                                            cmdCook(i).Left = frmRecipe.Width / 2
                                            Select Case i
                                                Case Is > 13
                                                    cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                                    cmdCook(i).top = cmdCook(i - 7).top
                                                Case 13
                                                    cmdCook(i).Left = cmdCook(0).Width
                                                    cmdCook(i).top = cmdCook(6).top
                                                Case Else
                                                    cmdCook(i).Width = frmRecipe.Width / 2
                                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                            End Select
                                            If i > 13 Then
                                                For b = 0 To i
                                                    cmdCook(b).FontTextCaption.Size = 11
                                                Next b
                                            Else
                                                For b = 0 To i
                                                    cmdCook(b).FontTextCaption.Size = 14
                                                Next b
                                            End If
                                        Case Is > 6
                                            For b = 0 To 6
                                                cmdCook(b).Width = frmRecipe.Width / 2
                                            Next b
                                            If i > cmdCook.UBound Then Load cmdCook(i)
                                            cmdCook(i).Visible = True
                                            cmdCook(i).Left = frmRecipe.Width / 2
                                            cmdCook(i).Width = frmRecipe.Width / 2
                                            cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                    End Select
                                    cmdCook(i).Visible = True
                                    If rs.Fields("Line_Code") = 0 Then
                                        cmdCook(i).Caption = UCase(rs.Fields("Description"))
                                    Else
                                        If rs.Fields("Line_Type") = 6 Or rs.Fields("Line_Type") = 7 Or rs.Fields("Line_Type") = 3 Then
                                            ActiveReadServer1 "Select Short_Description from Products where Product_Code='" & rs.Fields("Line_Code") & "'"
                                            If rs1.RecordCount > 0 Then Description = UCase(rs1.Fields("Short_Description") & "")
                                            rs1.Close
                                        Else
                                            Description = UCase(Trim(Mid(rs.Fields("Description"), 1, InStrRev(rs.Fields("Description"), ",") - 1)))
                                            If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                                        End If
                                        cmdCook(i).Caption = UCase(Description)
                                    End If
                                    cmdCook(i).Tag = rs.Fields("Line_Code")
                                    rs.MoveNext
                                Wend
                                rs.Close
                            End If
                    End Select
                End If
            Case 2
                If frmBar.grdMenu.Rows > 0 Then
                    Select Case frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 2)
                        Case 0, 8
                            i = i + 1
                            Select Case i
                                Case Is > 12
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 3
                                    Next b
                                    For b = 7 To 12
                                        cmdCook(b).Width = frmRecipe.Width / 3
                                        cmdCook(b).Left = cmdCook(0).Width
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    Select Case i
                                        Case Is > 13
                                            cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                            cmdCook(i).top = cmdCook(i - 7).top
                                        Case 13
                                            cmdCook(i).Left = cmdCook(0).Width
                                            cmdCook(i).top = cmdCook(6).top
                                        Case Else
                                            cmdCook(i).Width = frmRecipe.Width / 2
                                            cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                    End Select
                                    If i > 13 Then
                                        For b = 0 To i
                                            cmdCook(b).FontTextCaption.Size = 11
                                        Next b
                                    Else
                                        For b = 0 To i
                                            cmdCook(b).FontTextCaption.Size = 14
                                        Next b
                                    End If
                                Case Is > 6
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 2
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    cmdCook(i).Width = frmRecipe.Width / 2
                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                            End Select
                            cmdCook(i).Visible = True
                            If frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 2) = 8 Then
                                cmdCook(i).Caption = "EXIT"
                                cmdCook(i).Tag = "Exit"
                                cmdCook(i).BackColor = &H80C0FF
                            Else
                                cmdCook(i).Caption = UCase(frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 3))
                                cmdCook(i).Tag = frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 1)
                            End If
                            If frmBar.grdMenu.Row <> frmBar.grdMenu.Rows - 1 Then
                                frmBar.grdMenu.Row = frmBar.grdMenu.Row + 1
                                GoTo top
                            End If
                        Case 2, 3, 6, 7
                            i = i + 1
                            Select Case i
                                Case Is > 12
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 3
                                    Next b
                                    For b = 7 To 12
                                        cmdCook(b).Width = frmRecipe.Width / 3
                                        cmdCook(b).Left = cmdCook(0).Width
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    Select Case i
                                        Case Is > 13
                                            cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                            cmdCook(i).top = cmdCook(i - 7).top
                                        Case 13
                                            cmdCook(i).Left = cmdCook(0).Width
                                            cmdCook(i).top = cmdCook(6).top
                                        Case Else
                                            cmdCook(i).Width = frmRecipe.Width / 2
                                            cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                    End Select
                                    If i > 13 Then
                                        For b = 0 To i
                                            cmdCook(b).FontTextCaption.Size = 11
                                        Next b
                                    Else
                                        For b = 0 To i
                                            cmdCook(b).FontTextCaption.Size = 14
                                        Next b
                                    End If
                                Case Is > 6
                                    For b = 0 To 6
                                        cmdCook(b).Width = frmRecipe.Width / 2
                                    Next b
                                    If i > cmdCook.UBound Then Load cmdCook(i)
                                    cmdCook(i).Visible = True
                                    cmdCook(i).Left = frmRecipe.Width / 2
                                    cmdCook(i).Width = frmRecipe.Width / 2
                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                            End Select
                            cmdCook(i).Visible = True
                            If frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 2) = 6 Or frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 2) = 7 Or frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 2) = 3 Then
                                ActiveReadServer1 "Select Short_Description from Products where Product_Code='" & frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 1) & "'"
                                If rs1.RecordCount > 0 Then Description = rs1.Fields("Short_Description") & ""
                                rs1.Close
                            Else
                                Description = Trim(Mid(frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 3), 1, InStrRev(frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 3), ",") - 1))
                                If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                            End If
                            If frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 2) = 6 Or frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 2) = 7 Then
                                cmdCook(i).Caption = UCase(Description & Chr(160))
                            Else
                                cmdCook(i).Caption = UCase(Description)
                            End If
                            cmdCook(i).Tag = frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 1)
                            If frmBar.grdMenu.Row <> frmBar.grdMenu.Rows - 1 Then
                                frmBar.grdMenu.Row = frmBar.grdMenu.Row + 1
                                GoTo top
                            End If
                        Case 1
                            If i = -1 Then
                                ActiveReadServer "Select * from recipes where Product_Code= '" & frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 1) & "' order by Line_No"
                                frmBar.lblKeyRegister.TextAlign = fmTextAlignLeft
                                frmBar.lblKeyRegister = Mid(frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 3), 1, InStrRev(frmBar.grdMenu.TextMatrix(frmBar.grdMenu.Row, 3), ",") - 1)
                                While Not rs.EOF
                                    i = i + 1
                                    Select Case i
                                        Case Is > 12
                                            For b = 0 To 6
                                                cmdCook(b).Width = frmRecipe.Width / 3
                                            Next b
                                            For b = 7 To 12
                                                cmdCook(b).Width = frmRecipe.Width / 3
                                                cmdCook(b).Left = cmdCook(0).Width
                                            Next b
                                            If i > cmdCook.UBound Then Load cmdCook(i)
                                            cmdCook(i).Visible = True
                                            cmdCook(i).Left = frmRecipe.Width / 2
                                            Select Case i
                                                Case Is > 13
                                                    cmdCook(i).Left = (cmdCook(0).Width * 2) - 20
                                                    cmdCook(i).top = cmdCook(i - 7).top
                                                Case 13
                                                    cmdCook(i).Left = cmdCook(0).Width
                                                    cmdCook(i).top = cmdCook(6).top
                                                Case Else
                                                    cmdCook(i).Width = frmRecipe.Width / 2
                                                    cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                            End Select
                                            If i > 13 Then
                                                For b = 0 To i
                                                    cmdCook(b).FontTextCaption.Size = 11
                                                Next b
                                            Else
                                                For b = 0 To i
                                                    cmdCook(b).FontTextCaption.Size = 14
                                                Next b
                                            End If
                                        Case Is > 6
                                            For b = 0 To 6
                                                cmdCook(b).Width = frmRecipe.Width / 2
                                            Next b
                                            If i > cmdCook.UBound Then Load cmdCook(i)
                                            cmdCook(i).Visible = True
                                            cmdCook(i).Left = frmRecipe.Width / 2
                                            cmdCook(i).Width = frmRecipe.Width / 2
                                            cmdCook(i).top = (i - 7) * cmdCook(i).Height
                                    End Select
                                    cmdCook(i).Visible = True
                                    If rs.Fields("Line_Code") = 0 Then
                                        cmdCook(i).Caption = UCase(rs.Fields("Description"))
                                    Else
                                        Description = Trim(Mid(rs.Fields("Description"), 1, InStrRev(rs.Fields("Description"), ",") - 1))
                                        If Right(Description, 4) = "each" Then Description = Trim(Left(Description, Len(Description) - 4))
                                        cmdCook(i).Caption = UCase(Description)
                                    End If
                                    cmdCook(i).Tag = rs.Fields("Line_Code")
                                    rs.MoveNext
                                Wend
                                rs.Close
                            End If
                    End Select
                End If
        End Select
    End Select
    If cmdCook(0).Visible = False Then
        lblClick.Visible = True
    Else
        lblClick.Visible = False
    End If
End Sub
Private Sub Form_Click()
    If cmdCook(0).Visible = False Then
        Me.Hide
    End If
End Sub
Private Sub Form_Load()
    If frmSplash.Height < 10000 Then
        On Error Resume Next
        DoEvents
        Me.Width = Me.Width * 0.782
        Me.Height = Me.Height * 0.782
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.782
            Me.Controls(i).Left = Me.Controls(i).Left * 0.778
            Me.Controls(i).Height = Me.Controls(i).Height * 0.78
            Me.Controls(i).top = Me.Controls(i).top * 0.78
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
    
        On Error GoTo 0
    End If
    DoEvents
End Sub

Private Sub lblClick_Click()
    If cmdCook(0).Visible = False Then
        Me.Hide
    End If
End Sub

Private Sub Timer1_Timer()
 If cmdCook(0).Visible = False Then
 If frmRecipe.Visible Then
        Me.Hide
        End If
    End If
    Timer1.Enabled = False
End Sub
