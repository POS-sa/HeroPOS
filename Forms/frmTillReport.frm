VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmTillReport 
   BackColor       =   &H0083CEEF&
   ClientHeight    =   11400
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FBCFBF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmTillReport.frx":0000
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin btButtonEx.ButtonEx cmdNext 
      Height          =   810
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1429
      Appearance      =   3
      BackColor       =   2720171
      Caption         =   "7"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   26.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picHoldFocus 
      Height          =   675
      Left            =   480
      ScaleHeight     =   615
      ScaleWidth      =   405
      TabIndex        =   126
      Top             =   390
      Width           =   465
   End
   Begin VSFlex8Ctl.VSFlexGrid grdCash 
      Height          =   8790
      Left            =   60
      TabIndex        =   124
      Top             =   60
      Visible         =   0   'False
      Width           =   75
      _cx             =   132
      _cy             =   15505
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
   Begin VB.PictureBox picHist 
      BackColor       =   &H00E0F3FC&
      Height          =   7005
      Left            =   1590
      ScaleHeight     =   6945
      ScaleWidth      =   10965
      TabIndex        =   107
      Top             =   3600
      Width           =   11025
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   1
         Left            =   2760
         TabIndex        =   108
         Top             =   0
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524298
         _ExtentX        =   4842
         _ExtentY        =   2990
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":E61E
         textLT          =   "frmTillReport.frx":E636
         textCT          =   "frmTillReport.frx":E64E
         textRT          =   "frmTillReport.frx":E666
         textLM          =   "frmTillReport.frx":E67E
         textRM          =   "frmTillReport.frx":E696
         textLB          =   "frmTillReport.frx":E6AE
         textCB          =   "frmTillReport.frx":E6C6
         textRB          =   "frmTillReport.frx":E6DE
         colorBack       =   "frmTillReport.frx":E6F6
         colorIntern     =   "frmTillReport.frx":E720
         colorMO         =   "frmTillReport.frx":E74A
         colorFocus      =   "frmTillReport.frx":E774
         colorDisabled   =   "frmTillReport.frx":E79E
         colorPressed    =   "frmTillReport.frx":E7C8
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   0
         Left            =   0
         TabIndex        =   109
         Top             =   0
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524298
         _ExtentX        =   4842
         _ExtentY        =   2990
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
            Size            =   11.25
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
         Shape           =   1
         CornerFactor    =   10
         Surface         =   1
         BackColorContainer=   3119822
         SpecialEffect   =   1
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":E7F2
         textLT          =   "frmTillReport.frx":E80A
         textCT          =   "frmTillReport.frx":E822
         textRT          =   "frmTillReport.frx":E83A
         textLM          =   "frmTillReport.frx":E852
         textRM          =   "frmTillReport.frx":E86A
         textLB          =   "frmTillReport.frx":E882
         textCB          =   "frmTillReport.frx":E89A
         textRB          =   "frmTillReport.frx":E8B2
         colorBack       =   "frmTillReport.frx":E8CA
         colorIntern     =   "frmTillReport.frx":E8F4
         colorMO         =   "frmTillReport.frx":E91E
         colorFocus      =   "frmTillReport.frx":E948
         colorDisabled   =   "frmTillReport.frx":E972
         colorPressed    =   "frmTillReport.frx":E99C
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   2
         Left            =   5520
         TabIndex        =   110
         Top             =   0
         Visible         =   0   'False
         Width           =   2700
         _Version        =   524298
         _ExtentX        =   4762
         _ExtentY        =   2990
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":E9C6
         textLT          =   "frmTillReport.frx":E9DE
         textCT          =   "frmTillReport.frx":E9F6
         textRT          =   "frmTillReport.frx":EA0E
         textLM          =   "frmTillReport.frx":EA26
         textRM          =   "frmTillReport.frx":EA3E
         textLB          =   "frmTillReport.frx":EA56
         textCB          =   "frmTillReport.frx":EA6E
         textRB          =   "frmTillReport.frx":EA86
         colorBack       =   "frmTillReport.frx":EA9E
         colorIntern     =   "frmTillReport.frx":EAC8
         colorMO         =   "frmTillReport.frx":EAF2
         colorFocus      =   "frmTillReport.frx":EB1C
         colorDisabled   =   "frmTillReport.frx":EB46
         colorPressed    =   "frmTillReport.frx":EB70
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   3
         Left            =   8235
         TabIndex        =   111
         Top             =   0
         Visible         =   0   'False
         Width           =   2740
         _Version        =   524298
         _ExtentX        =   4833
         _ExtentY        =   2990
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
            Size            =   11.25
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
         Shape           =   1
         CornerFactor    =   10
         Surface         =   1
         BackColorContainer=   3119822
         SpecialEffect   =   1
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":EB9A
         textLT          =   "frmTillReport.frx":EBB2
         textCT          =   "frmTillReport.frx":EBCA
         textRT          =   "frmTillReport.frx":EBE2
         textLM          =   "frmTillReport.frx":EBFA
         textRM          =   "frmTillReport.frx":EC12
         textLB          =   "frmTillReport.frx":EC2A
         textCB          =   "frmTillReport.frx":EC42
         textRB          =   "frmTillReport.frx":EC5A
         colorBack       =   "frmTillReport.frx":EC72
         colorIntern     =   "frmTillReport.frx":EC9C
         colorMO         =   "frmTillReport.frx":ECC6
         colorFocus      =   "frmTillReport.frx":ECF0
         colorDisabled   =   "frmTillReport.frx":ED1A
         colorPressed    =   "frmTillReport.frx":ED44
         Orientation     =   6
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1705
         Index           =   4
         Left            =   0
         TabIndex        =   112
         Top             =   1710
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524298
         _ExtentX        =   4842
         _ExtentY        =   3007
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":ED6E
         textLT          =   "frmTillReport.frx":ED86
         textCT          =   "frmTillReport.frx":ED9E
         textRT          =   "frmTillReport.frx":EDB6
         textLM          =   "frmTillReport.frx":EDCE
         textRM          =   "frmTillReport.frx":EDE6
         textLB          =   "frmTillReport.frx":EDFE
         textCB          =   "frmTillReport.frx":EE16
         textRB          =   "frmTillReport.frx":EE2E
         colorBack       =   "frmTillReport.frx":EE46
         colorIntern     =   "frmTillReport.frx":EE70
         colorMO         =   "frmTillReport.frx":EE9A
         colorFocus      =   "frmTillReport.frx":EEC4
         colorDisabled   =   "frmTillReport.frx":EEEE
         colorPressed    =   "frmTillReport.frx":EF18
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1705
         Index           =   5
         Left            =   2760
         TabIndex        =   113
         Top             =   1710
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524298
         _ExtentX        =   4842
         _ExtentY        =   3007
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":EF42
         textLT          =   "frmTillReport.frx":EF5A
         textCT          =   "frmTillReport.frx":EF72
         textRT          =   "frmTillReport.frx":EF8A
         textLM          =   "frmTillReport.frx":EFA2
         textRM          =   "frmTillReport.frx":EFBA
         textLB          =   "frmTillReport.frx":EFD2
         textCB          =   "frmTillReport.frx":EFEA
         textRB          =   "frmTillReport.frx":F002
         colorBack       =   "frmTillReport.frx":F01A
         colorIntern     =   "frmTillReport.frx":F044
         colorMO         =   "frmTillReport.frx":F06E
         colorFocus      =   "frmTillReport.frx":F098
         colorDisabled   =   "frmTillReport.frx":F0C2
         colorPressed    =   "frmTillReport.frx":F0EC
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   6
         Left            =   5520
         TabIndex        =   114
         Top             =   1725
         Visible         =   0   'False
         Width           =   2700
         _Version        =   524298
         _ExtentX        =   4762
         _ExtentY        =   2990
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":F116
         textLT          =   "frmTillReport.frx":F12E
         textCT          =   "frmTillReport.frx":F146
         textRT          =   "frmTillReport.frx":F15E
         textLM          =   "frmTillReport.frx":F176
         textRM          =   "frmTillReport.frx":F18E
         textLB          =   "frmTillReport.frx":F1A6
         textCB          =   "frmTillReport.frx":F1BE
         textRB          =   "frmTillReport.frx":F1D6
         colorBack       =   "frmTillReport.frx":F1EE
         colorIntern     =   "frmTillReport.frx":F218
         colorMO         =   "frmTillReport.frx":F242
         colorFocus      =   "frmTillReport.frx":F26C
         colorDisabled   =   "frmTillReport.frx":F296
         colorPressed    =   "frmTillReport.frx":F2C0
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   7
         Left            =   8235
         TabIndex        =   115
         Top             =   1725
         Visible         =   0   'False
         Width           =   2740
         _Version        =   524298
         _ExtentX        =   4833
         _ExtentY        =   2990
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":F2EA
         textLT          =   "frmTillReport.frx":F302
         textCT          =   "frmTillReport.frx":F31A
         textRT          =   "frmTillReport.frx":F332
         textLM          =   "frmTillReport.frx":F34A
         textRM          =   "frmTillReport.frx":F362
         textLB          =   "frmTillReport.frx":F37A
         textCB          =   "frmTillReport.frx":F392
         textRB          =   "frmTillReport.frx":F3AA
         colorBack       =   "frmTillReport.frx":F3C2
         colorIntern     =   "frmTillReport.frx":F3EC
         colorMO         =   "frmTillReport.frx":F416
         colorFocus      =   "frmTillReport.frx":F440
         colorDisabled   =   "frmTillReport.frx":F46A
         colorPressed    =   "frmTillReport.frx":F494
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1785
         Index           =   12
         Left            =   0
         TabIndex        =   116
         Top             =   5160
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524298
         _ExtentX        =   4842
         _ExtentY        =   3149
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
            Size            =   11.25
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
         Shape           =   1
         CornerFactor    =   10
         Surface         =   1
         BackColorContainer=   3119822
         SpecialEffect   =   1
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":F4BE
         textLT          =   "frmTillReport.frx":F4D6
         textCT          =   "frmTillReport.frx":F4EE
         textRT          =   "frmTillReport.frx":F506
         textLM          =   "frmTillReport.frx":F51E
         textRM          =   "frmTillReport.frx":F536
         textLB          =   "frmTillReport.frx":F54E
         textCB          =   "frmTillReport.frx":F566
         textRB          =   "frmTillReport.frx":F57E
         colorBack       =   "frmTillReport.frx":F596
         colorIntern     =   "frmTillReport.frx":F5C0
         colorMO         =   "frmTillReport.frx":F5EA
         colorFocus      =   "frmTillReport.frx":F614
         colorDisabled   =   "frmTillReport.frx":F63E
         colorPressed    =   "frmTillReport.frx":F668
         Orientation     =   8
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1785
         Index           =   13
         Left            =   2760
         TabIndex        =   117
         Top             =   5175
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524298
         _ExtentX        =   4842
         _ExtentY        =   3149
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":F692
         textLT          =   "frmTillReport.frx":F6AA
         textCT          =   "frmTillReport.frx":F6C2
         textRT          =   "frmTillReport.frx":F6DA
         textLM          =   "frmTillReport.frx":F6F2
         textRM          =   "frmTillReport.frx":F70A
         textLB          =   "frmTillReport.frx":F722
         textCB          =   "frmTillReport.frx":F73A
         textRB          =   "frmTillReport.frx":F752
         colorBack       =   "frmTillReport.frx":F76A
         colorIntern     =   "frmTillReport.frx":F794
         colorMO         =   "frmTillReport.frx":F7BE
         colorFocus      =   "frmTillReport.frx":F7E8
         colorDisabled   =   "frmTillReport.frx":F812
         colorPressed    =   "frmTillReport.frx":F83C
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1785
         Index           =   14
         Left            =   5520
         TabIndex        =   118
         Top             =   5175
         Visible         =   0   'False
         Width           =   2700
         _Version        =   524298
         _ExtentX        =   4762
         _ExtentY        =   3149
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":F866
         textLT          =   "frmTillReport.frx":F87E
         textCT          =   "frmTillReport.frx":F896
         textRT          =   "frmTillReport.frx":F8AE
         textLM          =   "frmTillReport.frx":F8C6
         textRM          =   "frmTillReport.frx":F8DE
         textLB          =   "frmTillReport.frx":F8F6
         textCB          =   "frmTillReport.frx":F90E
         textRB          =   "frmTillReport.frx":F926
         colorBack       =   "frmTillReport.frx":F93E
         colorIntern     =   "frmTillReport.frx":F968
         colorMO         =   "frmTillReport.frx":F992
         colorFocus      =   "frmTillReport.frx":F9BC
         colorDisabled   =   "frmTillReport.frx":F9E6
         colorPressed    =   "frmTillReport.frx":FA10
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1785
         Index           =   15
         Left            =   8235
         TabIndex        =   119
         Top             =   5175
         Visible         =   0   'False
         Width           =   2740
         _Version        =   524298
         _ExtentX        =   4833
         _ExtentY        =   3149
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
            Size            =   11.25
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
         Shape           =   1
         CornerFactor    =   10
         Surface         =   1
         BackColorContainer=   3119822
         SpecialEffect   =   1
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":FA3A
         textLT          =   "frmTillReport.frx":FA52
         textCT          =   "frmTillReport.frx":FA6A
         textRT          =   "frmTillReport.frx":FA82
         textLM          =   "frmTillReport.frx":FA9A
         textRM          =   "frmTillReport.frx":FAB2
         textLB          =   "frmTillReport.frx":FACA
         textCB          =   "frmTillReport.frx":FAE2
         textRB          =   "frmTillReport.frx":FAFA
         colorBack       =   "frmTillReport.frx":FB12
         colorIntern     =   "frmTillReport.frx":FB3C
         colorMO         =   "frmTillReport.frx":FB66
         colorFocus      =   "frmTillReport.frx":FB90
         colorDisabled   =   "frmTillReport.frx":FBBA
         colorPressed    =   "frmTillReport.frx":FBE4
         Orientation     =   7
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   8
         Left            =   0
         TabIndex        =   120
         Top             =   3450
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524298
         _ExtentX        =   4842
         _ExtentY        =   2990
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":FC0E
         textLT          =   "frmTillReport.frx":FC26
         textCT          =   "frmTillReport.frx":FC3E
         textRT          =   "frmTillReport.frx":FC56
         textLM          =   "frmTillReport.frx":FC6E
         textRM          =   "frmTillReport.frx":FC86
         textLB          =   "frmTillReport.frx":FC9E
         textCB          =   "frmTillReport.frx":FCB6
         textRB          =   "frmTillReport.frx":FCCE
         colorBack       =   "frmTillReport.frx":FCE6
         colorIntern     =   "frmTillReport.frx":FD10
         colorMO         =   "frmTillReport.frx":FD3A
         colorFocus      =   "frmTillReport.frx":FD64
         colorDisabled   =   "frmTillReport.frx":FD8E
         colorPressed    =   "frmTillReport.frx":FDB8
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   9
         Left            =   2760
         TabIndex        =   121
         Top             =   3450
         Visible         =   0   'False
         Width           =   2745
         _Version        =   524298
         _ExtentX        =   4842
         _ExtentY        =   2990
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":FDE2
         textLT          =   "frmTillReport.frx":FDFA
         textCT          =   "frmTillReport.frx":FE12
         textRT          =   "frmTillReport.frx":FE2A
         textLM          =   "frmTillReport.frx":FE42
         textRM          =   "frmTillReport.frx":FE5A
         textLB          =   "frmTillReport.frx":FE72
         textCB          =   "frmTillReport.frx":FE8A
         textRB          =   "frmTillReport.frx":FEA2
         colorBack       =   "frmTillReport.frx":FEBA
         colorIntern     =   "frmTillReport.frx":FEE4
         colorMO         =   "frmTillReport.frx":FF0E
         colorFocus      =   "frmTillReport.frx":FF38
         colorDisabled   =   "frmTillReport.frx":FF62
         colorPressed    =   "frmTillReport.frx":FF8C
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   10
         Left            =   5520
         TabIndex        =   122
         Top             =   3450
         Visible         =   0   'False
         Width           =   2700
         _Version        =   524298
         _ExtentX        =   4762
         _ExtentY        =   2990
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":FFB6
         textLT          =   "frmTillReport.frx":FFCE
         textCT          =   "frmTillReport.frx":FFE6
         textRT          =   "frmTillReport.frx":FFFE
         textLM          =   "frmTillReport.frx":10016
         textRM          =   "frmTillReport.frx":1002E
         textLB          =   "frmTillReport.frx":10046
         textCB          =   "frmTillReport.frx":1005E
         textRB          =   "frmTillReport.frx":10076
         colorBack       =   "frmTillReport.frx":1008E
         colorIntern     =   "frmTillReport.frx":100B8
         colorMO         =   "frmTillReport.frx":100E2
         colorFocus      =   "frmTillReport.frx":1010C
         colorDisabled   =   "frmTillReport.frx":10136
         colorPressed    =   "frmTillReport.frx":10160
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
      Begin BTNENHLib4.BtnEnh cmdCash 
         Height          =   1695
         Index           =   11
         Left            =   8235
         TabIndex        =   123
         Top             =   3450
         Visible         =   0   'False
         Width           =   2740
         _Version        =   524298
         _ExtentX        =   4833
         _ExtentY        =   2990
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
            Size            =   11.25
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
         CaptionWordWrapPerc=   100
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmTillReport.frx":1018A
         textLT          =   "frmTillReport.frx":101A2
         textCT          =   "frmTillReport.frx":101BA
         textRT          =   "frmTillReport.frx":101D2
         textLM          =   "frmTillReport.frx":101EA
         textRM          =   "frmTillReport.frx":10202
         textLB          =   "frmTillReport.frx":1021A
         textCB          =   "frmTillReport.frx":10232
         textRB          =   "frmTillReport.frx":1024A
         colorBack       =   "frmTillReport.frx":10262
         colorIntern     =   "frmTillReport.frx":1028C
         colorMO         =   "frmTillReport.frx":102B6
         colorFocus      =   "frmTillReport.frx":102E0
         colorDisabled   =   "frmTillReport.frx":1030A
         colorPressed    =   "frmTillReport.frx":10334
         Orientation     =   5
         UseAntialias    =   0   'False
         HollowFrame     =   -1  'True
         LightDirection  =   1
      End
   End
   Begin BTNENHLib4.BtnEnh cmdHist 
      Height          =   795
      Left            =   1440
      TabIndex        =   106
      Top             =   330
      Visible         =   0   'False
      Width           =   5025
      _Version        =   524298
      _ExtentX        =   8864
      _ExtentY        =   1402
      _StockProps     =   66
      Caption         =   "Click for active Cashup's"
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
      textCaption     =   "frmTillReport.frx":1035E
      textLT          =   "frmTillReport.frx":103F0
      textCT          =   "frmTillReport.frx":10408
      textRT          =   "frmTillReport.frx":10420
      textLM          =   "frmTillReport.frx":10438
      textRM          =   "frmTillReport.frx":10450
      textLB          =   "frmTillReport.frx":10468
      textCB          =   "frmTillReport.frx":10480
      textRB          =   "frmTillReport.frx":10498
      colorBack       =   "frmTillReport.frx":104B0
      colorIntern     =   "frmTillReport.frx":104DA
      colorMO         =   "frmTillReport.frx":10504
      colorFocus      =   "frmTillReport.frx":1052E
      colorDisabled   =   "frmTillReport.frx":10558
      colorPressed    =   "frmTillReport.frx":10582
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7770
      Top             =   1590
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4350
      Top             =   10740
   End
   Begin BTNENHLib4.BtnEnh cmdUser 
      Height          =   825
      Left            =   1440
      TabIndex        =   1
      Top             =   330
      Width           =   4155
      _Version        =   524298
      _ExtentX        =   7329
      _ExtentY        =   1455
      _StockProps     =   66
      Caption         =   "<None>"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   17.25
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
      BackColorContainer=   3119822
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
      textCaption     =   "frmTillReport.frx":105AC
      textLT          =   "frmTillReport.frx":10618
      textCT          =   "frmTillReport.frx":10630
      textRT          =   "frmTillReport.frx":10648
      textLM          =   "frmTillReport.frx":10660
      textRM          =   "frmTillReport.frx":10678
      textLB          =   "frmTillReport.frx":10690
      textCB          =   "frmTillReport.frx":106A8
      textRB          =   "frmTillReport.frx":106C0
      colorBack       =   "frmTillReport.frx":106D8
      colorIntern     =   "frmTillReport.frx":10702
      colorMO         =   "frmTillReport.frx":1072C
      colorFocus      =   "frmTillReport.frx":10756
      colorDisabled   =   "frmTillReport.frx":10780
      colorPressed    =   "frmTillReport.frx":107AA
      Orientation     =   1
      TextCaptionAlignment=   0
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh BtnEnh1 
      Height          =   825
      Index           =   3
      Left            =   5580
      TabIndex        =   2
      Top             =   330
      Width           =   855
      _Version        =   524298
      _ExtentX        =   1508
      _ExtentY        =   1455
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
      BackColorContainer=   3119822
      ButtonRaiseFactor=   3
      SmoothEdges     =   1
      SpecialEffect   =   2
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      FlatPillowFactor=   3
      UserData        =   0.1
      textCaption     =   "frmTillReport.frx":107D4
      textLT          =   "frmTillReport.frx":10836
      textCT          =   "frmTillReport.frx":1084E
      textRT          =   "frmTillReport.frx":10866
      textLM          =   "frmTillReport.frx":1087E
      textRM          =   "frmTillReport.frx":10896
      textLB          =   "frmTillReport.frx":108AE
      textCB          =   "frmTillReport.frx":108C6
      textRB          =   "frmTillReport.frx":108DE
      colorBack       =   "frmTillReport.frx":108F6
      colorIntern     =   "frmTillReport.frx":10920
      colorMO         =   "frmTillReport.frx":1094A
      colorFocus      =   "frmTillReport.frx":10974
      colorDisabled   =   "frmTillReport.frx":1099E
      colorPressed    =   "frmTillReport.frx":109C8
      Orientation     =   3
      TextCaptionAlignment=   0
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1875
      Index           =   2
      Left            =   390
      TabIndex        =   18
      Top             =   3630
      Width           =   1065
      _Version        =   524298
      _ExtentX        =   1879
      _ExtentY        =   3307
      _StockProps     =   66
      Caption         =   "Print Count Slip"
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
      BackColorContainer=   14737632
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmTillReport.frx":109F2
      textLT          =   "frmTillReport.frx":10A72
      textCT          =   "frmTillReport.frx":10A8A
      textRT          =   "frmTillReport.frx":10AA2
      textLM          =   "frmTillReport.frx":10ABA
      textRM          =   "frmTillReport.frx":10AD2
      textLB          =   "frmTillReport.frx":10AEA
      textCB          =   "frmTillReport.frx":10B02
      textRB          =   "frmTillReport.frx":10B1A
      colorBack       =   "frmTillReport.frx":10B32
      colorIntern     =   "frmTillReport.frx":10B5C
      colorMO         =   "frmTillReport.frx":10B86
      colorFocus      =   "frmTillReport.frx":10BB0
      colorDisabled   =   "frmTillReport.frx":10BDA
      colorPressed    =   "frmTillReport.frx":10C04
      Orientation     =   2
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1875
      Index           =   0
      Left            =   390
      TabIndex        =   19
      Top             =   5505
      Width           =   1065
      _Version        =   524298
      _ExtentX        =   1879
      _ExtentY        =   3307
      _StockProps     =   66
      Caption         =   "Capture Cashup"
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
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmTillReport.frx":10C2E
      textLT          =   "frmTillReport.frx":10CAA
      textCT          =   "frmTillReport.frx":10CC2
      textRT          =   "frmTillReport.frx":10CDA
      textLM          =   "frmTillReport.frx":10CF2
      textRM          =   "frmTillReport.frx":10D0A
      textLB          =   "frmTillReport.frx":10D22
      textCB          =   "frmTillReport.frx":10D3A
      textRB          =   "frmTillReport.frx":10D52
      colorBack       =   "frmTillReport.frx":10D6A
      colorIntern     =   "frmTillReport.frx":10D94
      colorMO         =   "frmTillReport.frx":10DBE
      colorFocus      =   "frmTillReport.frx":10DE8
      colorDisabled   =   "frmTillReport.frx":10E12
      colorPressed    =   "frmTillReport.frx":10E3C
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1935
      Index           =   1
      Left            =   390
      TabIndex        =   20
      Top             =   7380
      Width           =   1065
      _Version        =   524298
      _ExtentX        =   1879
      _ExtentY        =   3413
      _StockProps     =   66
      Caption         =   "Print Cashup"
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
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmTillReport.frx":10E66
      textLT          =   "frmTillReport.frx":10EDE
      textCT          =   "frmTillReport.frx":10EF6
      textRT          =   "frmTillReport.frx":10F0E
      textLM          =   "frmTillReport.frx":10F26
      textRM          =   "frmTillReport.frx":10F3E
      textLB          =   "frmTillReport.frx":10F56
      textCB          =   "frmTillReport.frx":10F6E
      textRB          =   "frmTillReport.frx":10F86
      colorBack       =   "frmTillReport.frx":10F9E
      colorIntern     =   "frmTillReport.frx":10FC8
      colorMO         =   "frmTillReport.frx":10FF2
      colorFocus      =   "frmTillReport.frx":1101C
      colorDisabled   =   "frmTillReport.frx":11046
      colorPressed    =   "frmTillReport.frx":11070
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1905
      Index           =   3
      Left            =   390
      TabIndex        =   21
      Top             =   9315
      Width           =   1065
      _Version        =   524298
      _ExtentX        =   1879
      _ExtentY        =   3360
      _StockProps     =   66
      Caption         =   "Finalize Cashup"
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
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmTillReport.frx":1109A
      textLT          =   "frmTillReport.frx":11118
      textCT          =   "frmTillReport.frx":11130
      textRT          =   "frmTillReport.frx":11148
      textLM          =   "frmTillReport.frx":11160
      textRM          =   "frmTillReport.frx":11178
      textLB          =   "frmTillReport.frx":11190
      textCB          =   "frmTillReport.frx":111A8
      textRB          =   "frmTillReport.frx":111C0
      colorBack       =   "frmTillReport.frx":111D8
      colorIntern     =   "frmTillReport.frx":11202
      colorMO         =   "frmTillReport.frx":1122C
      colorFocus      =   "frmTillReport.frx":11256
      colorDisabled   =   "frmTillReport.frx":11280
      colorPressed    =   "frmTillReport.frx":112AA
      Orientation     =   4
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   21
      Left            =   12930
      TabIndex        =   22
      Top             =   4590
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "Wed"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   22
      Left            =   12930
      TabIndex        =   23
      Top             =   5490
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "Thu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   24
      Left            =   12930
      TabIndex        =   24
      Top             =   7290
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "Sat"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   25
      Left            =   12930
      TabIndex        =   25
      Top             =   8190
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "Sun"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   6
      Left            =   12930
      TabIndex        =   26
      Top             =   9090
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   3119822
      Caption         =   "6"
      CaptionOffsetX  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   23
      Left            =   12930
      TabIndex        =   27
      Top             =   6390
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "Fri"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   19
      Left            =   12930
      TabIndex        =   28
      Top             =   2790
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "Mon"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   20
      Left            =   12930
      TabIndex        =   29
      Top             =   3690
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "Tue"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   5
      Left            =   12930
      TabIndex        =   30
      Top             =   1890
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   3119822
      Caption         =   "5"
      CaptionOffsetX  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   30
      Left            =   1440
      TabIndex        =   77
      Top             =   1140
      Visible         =   0   'False
      Width           =   5025
      _cx             =   8864
      _cy             =   53
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   6534620
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   650
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
      WallPaper       =   "frmTillReport.frx":112D4
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   795
      Left            =   6570
      TabIndex        =   103
      Top             =   330
      Visible         =   0   'False
      Width           =   8565
      _Version        =   524298
      _ExtentX        =   15108
      _ExtentY        =   1402
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
      textCaption     =   "frmTillReport.frx":13177
      textLT          =   "frmTillReport.frx":131FD
      textCT          =   "frmTillReport.frx":13215
      textRT          =   "frmTillReport.frx":1322D
      textLM          =   "frmTillReport.frx":13245
      textRM          =   "frmTillReport.frx":1325D
      textLB          =   "frmTillReport.frx":13275
      textCB          =   "frmTillReport.frx":1328D
      textRB          =   "frmTillReport.frx":132A5
      colorBack       =   "frmTillReport.frx":132BD
      colorIntern     =   "frmTillReport.frx":132E7
      colorMO         =   "frmTillReport.frx":13311
      colorFocus      =   "frmTillReport.frx":1333B
      colorDisabled   =   "frmTillReport.frx":13365
      colorPressed    =   "frmTillReport.frx":1338F
   End
   Begin BTNENHLib4.BtnEnh fmClock 
      Height          =   1455
      Left            =   4170
      TabIndex        =   105
      Top             =   1890
      Visible         =   0   'False
      Width           =   3165
      _Version        =   524298
      _ExtentX        =   5583
      _ExtentY        =   2566
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
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
      Shape           =   1
      CornerFactor    =   10
      Surface         =   3
      BackColorContainer=   16777215
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmTillReport.frx":133B9
      textLT          =   "frmTillReport.frx":133D1
      textCT          =   "frmTillReport.frx":133E9
      textRT          =   "frmTillReport.frx":13401
      textLM          =   "frmTillReport.frx":13419
      textRM          =   "frmTillReport.frx":13431
      textLB          =   "frmTillReport.frx":13449
      textCB          =   "frmTillReport.frx":13461
      textRB          =   "frmTillReport.frx":13479
      colorBack       =   "frmTillReport.frx":13491
      colorIntern     =   "frmTillReport.frx":134BB
      colorMO         =   "frmTillReport.frx":134E5
      colorFocus      =   "frmTillReport.frx":1350F
      colorDisabled   =   "frmTillReport.frx":13539
      colorPressed    =   "frmTillReport.frx":13563
      HollowFrame     =   -1  'True
      LightDirection  =   8
   End
   Begin VB.Label lblClockinTime 
      Height          =   405
      Left            =   7530
      TabIndex        =   125
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label QAnswer 
      Height          =   915
      Left            =   5460
      TabIndex        =   104
      Top             =   2340
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   12
      Left            =   7620
      TabIndex        =   102
      Top             =   6780
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Service Charges:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   13
      Left            =   11070
      Top             =   6765
      Width           =   1365
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   19
      Left            =   9330
      TabIndex        =   101
      Top             =   6825
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   18
      Left            =   11100
      TabIndex        =   100
      Top             =   6825
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   12
      Left            =   9300
      Top             =   6765
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   24
      Left            =   1620
      TabIndex        =   99
      Top             =   7425
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "-Pickups:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   25
      Left            =   1620
      TabIndex        =   98
      Top             =   7920
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "+Deposits:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   26
      Left            =   1620
      TabIndex        =   97
      Top             =   8430
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "+Receive on Acc:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   38
      Left            =   5520
      Top             =   8415
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   39
      Left            =   5520
      Top             =   7920
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   40
      Left            =   5520
      Top             =   7425
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   41
      Left            =   5520
      Top             =   6930
      Width           =   1725
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   31
      Left            =   1620
      TabIndex        =   96
      Top             =   6840
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "-Payouts:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   39
      Left            =   3750
      TabIndex        =   95
      Top             =   6990
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   40
      Left            =   3750
      TabIndex        =   94
      Top             =   7485
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   41
      Left            =   3750
      TabIndex        =   93
      Top             =   7980
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   42
      Left            =   3750
      TabIndex        =   92
      Top             =   8475
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   43
      Left            =   5910
      TabIndex        =   91
      Top             =   6990
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   44
      Left            =   5910
      TabIndex        =   90
      Top             =   7485
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   45
      Left            =   5910
      TabIndex        =   89
      Top             =   7980
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   46
      Left            =   5910
      TabIndex        =   88
      Top             =   8475
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   37
      Left            =   3720
      Top             =   8415
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   36
      Left            =   3720
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   31
      Left            =   3720
      Top             =   7425
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   30
      Left            =   3720
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   32
      Left            =   9300
      Top             =   5265
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   33
      Left            =   9300
      Top             =   4770
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   34
      Left            =   9300
      Top             =   4275
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   35
      Left            =   9300
      Top             =   3780
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   46
      Left            =   9300
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   47
      Left            =   9300
      Top             =   6255
      Width           =   1695
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   17
      Left            =   11460
      TabIndex        =   87
      Top             =   3000
      Width           =   885
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "1561;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   11
      Left            =   11370
      Top             =   2940
      Width           =   1035
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   11
      Left            =   9300
      TabIndex        =   86
      Top             =   2970
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Power Off Count"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   27
      Left            =   2040
      TabIndex        =   85
      Top             =   9090
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Total Reported:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   405
      Index           =   42
      Left            =   3720
      Top             =   9030
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   33
      Left            =   2040
      TabIndex        =   84
      Top             =   9525
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Counted:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   405
      Index           =   43
      Left            =   3720
      Top             =   9525
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   34
      Left            =   2040
      TabIndex        =   83
      Top             =   9990
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Variance:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   405
      Index           =   50
      Left            =   3720
      Top             =   10020
      Width           =   1695
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   47
      Left            =   3750
      TabIndex        =   82
      Top             =   9090
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCash 
      Height          =   285
      Index           =   48
      Left            =   3750
      TabIndex        =   81
      Top             =   9600
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;503"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   49
      Left            =   3750
      TabIndex        =   80
      Top             =   10080
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCashdate 
      Height          =   285
      Left            =   8550
      TabIndex        =   79
      Top             =   10875
      Width           =   6375
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "11245;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   51
      Left            =   2730
      TabIndex        =   78
      Top             =   1470
      Width           =   4335
      ForeColor       =   0
      BackColor       =   16777215
      Size            =   "7646;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   35
      Left            =   9300
      TabIndex        =   76
      Top             =   2460
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "No Sale Count:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   51
      Left            =   11370
      Top             =   2430
      Width           =   1035
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   50
      Left            =   11460
      TabIndex        =   75
      Top             =   2490
      Width           =   885
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "1561;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   38
      Left            =   10830
      TabIndex        =   74
      Top             =   9270
      Width           =   1545
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2725;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   37
      Left            =   10830
      TabIndex        =   73
      Top             =   8760
      Width           =   1545
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2725;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   36
      Left            =   10830
      TabIndex        =   72
      Top             =   8280
      Width           =   1545
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2725;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   35
      Left            =   10830
      TabIndex        =   71
      Top             =   7770
      Width           =   1545
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2725;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   34
      Left            =   11100
      TabIndex        =   70
      Top             =   6315
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   33
      Left            =   11100
      TabIndex        =   69
      Top             =   5820
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   32
      Left            =   11100
      TabIndex        =   68
      Top             =   5325
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   31
      Left            =   11100
      TabIndex        =   67
      Top             =   4830
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   30
      Left            =   11100
      TabIndex        =   66
      Top             =   4335
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   29
      Left            =   11100
      TabIndex        =   65
      Top             =   3840
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   28
      Left            =   9330
      TabIndex        =   64
      Top             =   6315
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   27
      Left            =   9330
      TabIndex        =   63
      Top             =   5820
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   26
      Left            =   9330
      TabIndex        =   62
      Top             =   5325
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   25
      Left            =   9330
      TabIndex        =   61
      Top             =   4830
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   24
      Left            =   9330
      TabIndex        =   60
      Top             =   4335
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   23
      Left            =   9330
      TabIndex        =   59
      Top             =   3840
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2831;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   345
      Index           =   16
      Left            =   5880
      TabIndex        =   58
      Top             =   6360
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;609"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   15
      Left            =   3750
      TabIndex        =   57
      Top             =   6390
      Width           =   1575
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2778;556"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   14
      Left            =   5880
      TabIndex        =   56
      Top             =   5790
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   13
      Left            =   5880
      TabIndex        =   55
      Top             =   5280
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   12
      Left            =   5880
      TabIndex        =   54
      Top             =   4800
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   11
      Left            =   5880
      TabIndex        =   53
      Top             =   4320
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   10
      Left            =   5880
      TabIndex        =   52
      Top             =   3810
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   9
      Left            =   3750
      TabIndex        =   51
      Top             =   5790
      Width           =   1575
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2778;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   345
      Index           =   8
      Left            =   3750
      TabIndex        =   50
      Top             =   5280
      Width           =   1575
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2778;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   7
      Left            =   3750
      TabIndex        =   49
      Top             =   4800
      Width           =   1575
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2778;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   6
      Left            =   3750
      TabIndex        =   48
      Top             =   4320
      Width           =   1575
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2778;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   5
      Left            =   3750
      TabIndex        =   47
      Top             =   3810
      Width           =   1575
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "2778;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   4
      Left            =   11460
      TabIndex        =   46
      Top             =   2010
      Width           =   885
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "1561;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   3
      Left            =   11460
      TabIndex        =   45
      Top             =   1500
      Width           =   885
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "1561;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   2
      Left            =   2730
      TabIndex        =   44
      Top             =   2970
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0hrs 0min"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   1
      Left            =   2730
      TabIndex        =   43
      Top             =   2460
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Index           =   0
      Left            =   2730
      TabIndex        =   42
      Top             =   1980
      Width           =   1245
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2196;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   49
      Left            =   11070
      Top             =   5760
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   48
      Left            =   11070
      Top             =   6255
      Width           =   1365
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   32
      Left            =   7620
      TabIndex        =   41
      Top             =   5820
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Discount%:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   45
      Left            =   10740
      Top             =   9195
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   30
      Left            =   9060
      TabIndex        =   40
      Top             =   9225
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Tax Calculated:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   44
      Left            =   10740
      Top             =   8700
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   29
      Left            =   9060
      TabIndex        =   39
      Top             =   8730
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Tax Collected:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   25
      Left            =   10740
      Top             =   8205
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   28
      Left            =   8850
      TabIndex        =   38
      Top             =   8235
      Width           =   1815
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Non-Taxable Sales:"
      Size            =   "3201;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   24
      Left            =   10740
      Top             =   7710
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   18
      Left            =   9060
      TabIndex        =   37
      Top             =   7710
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Taxable Sales:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   19
      Left            =   7620
      TabIndex        =   36
      Top             =   6270
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Discount Value:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   23
      Left            =   7620
      TabIndex        =   35
      Top             =   3810
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Item Corrects:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   22
      Left            =   7620
      TabIndex        =   34
      Top             =   4305
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Voids:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   21
      Left            =   7620
      TabIndex        =   33
      Top             =   4800
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Returns:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   20
      Left            =   7620
      TabIndex        =   32
      Top             =   5295
      Width           =   1605
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Wastages:"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   29
      Left            =   11070
      Top             =   5265
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   28
      Left            =   11070
      Top             =   4770
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   27
      Left            =   11070
      Top             =   4275
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   26
      Left            =   11070
      Top             =   3780
      Width           =   1365
   End
   Begin MSForms.Label lblBDate 
      Height          =   285
      Left            =   12975
      TabIndex        =   31
      Top             =   10110
      Width           =   1875
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "12-01-2005"
      Size            =   "3307;503"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblTDate 
      Height          =   285
      Left            =   12960
      TabIndex        =   17
      Top             =   1560
      Width           =   1845
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "12-01-2005"
      Size            =   "3254;503"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   405
      Index           =   20
      Left            =   5520
      Top             =   6330
      Width           =   1725
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   14
      Left            =   1620
      TabIndex        =   16
      Top             =   6360
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Total:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   405
      Index           =   19
      Left            =   3720
      Top             =   6330
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   18
      Left            =   5520
      Top             =   3750
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   17
      Left            =   5520
      Top             =   4245
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   16
      Left            =   5520
      Top             =   4740
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   15
      Left            =   5520
      Top             =   5235
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   14
      Left            =   5520
      Top             =   5730
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   10
      Left            =   3720
      Top             =   5730
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   10
      Left            =   1620
      TabIndex        =   15
      Top             =   5760
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "+Loyalty Sales:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   9
      Left            =   3720
      Top             =   5235
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   9
      Left            =   1620
      TabIndex        =   14
      Top             =   5265
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "+Charge Sales:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   8
      Left            =   3720
      Top             =   4740
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   8
      Left            =   1620
      TabIndex        =   13
      Top             =   4770
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "+Voucher Sales:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   7
      Left            =   3720
      Top             =   4245
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   7
      Left            =   1620
      TabIndex        =   12
      Top             =   4275
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "+Card Sales:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   6
      Left            =   3720
      Top             =   3750
      Width           =   1695
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   6
      Left            =   1620
      TabIndex        =   11
      Top             =   3780
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "+Cash Sales:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   1890
      TabIndex        =   10
      Top             =   10875
      Width           =   5295
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "9340;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   540
      TabIndex        =   9
      Top             =   1440
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Date and Time:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   540
      TabIndex        =   8
      Top             =   1935
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Open Transaction No:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   540
      TabIndex        =   7
      Top             =   2430
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Close Transaction No:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   3
      Left            =   540
      TabIndex        =   6
      Top             =   2925
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Shift Duration:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   5
      Left            =   11370
      Top             =   1935
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   4
      Left            =   11370
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   3
      Left            =   2610
      Top             =   2895
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   2
      Left            =   2610
      Top             =   2400
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   1
      Left            =   2610
      Top             =   1905
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Index           =   0
      Left            =   2610
      Top             =   1410
      Width           =   4725
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   5
      Left            =   9300
      TabIndex        =   5
      Top             =   1965
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Customer Count:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Index           =   4
      Left            =   9270
      TabIndex        =   4
      Top             =   1470
      Width           =   2025
      ForeColor       =   0
      BackColor       =   16777215
      Caption         =   "Transaction Count:"
      Size            =   "3572;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCashupInfo 
      Height          =   615
      Left            =   6840
      TabIndex        =   3
      Top             =   540
      Width           =   7935
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "13996;1085"
      FontName        =   "Arial Narrow"
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image4 
      Height          =   345
      Left            =   12930
      Top             =   1500
      Width           =   1935
      BackColor       =   16777215
      Size            =   "3413;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Left            =   12930
      Top             =   10050
      Width           =   1935
      BackColor       =   16777215
      Size            =   "3413;609"
   End
   Begin MSForms.Image Image3 
      Height          =   9135
      Left            =   12780
      Top             =   1350
      Width           =   2235
      BackColor       =   14742524
      Size            =   "3942;16113"
   End
   Begin MSForms.Image newBack 
      Height          =   1785
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   285
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "503;3149"
   End
End
Attribute VB_Name = "frmTillReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LoadCashUp(CashupNo)
    On Error Resume Next
    If cmdHist.Visible = False Then
        ActiveReadServer "Select Function_Key,Date_Time from User_Journal where user_No= " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1) & " and line_No = " & _
        "(Select Max(Line_No) from User_Journal where function_Key in (3) and User_No=" & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1) & ")"
        If rs.RecordCount > 0 Then
            If rs.Fields("Function_Key").Value = 3 Then
                lblClockinTime.Caption = rs.Fields("Date_Time")
            End If
        End If
        rs.Close
        ActiveReadServer "Select * from Counters where User_no= " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1) & " and isnull(Finalized,0) = 0"
    Else
        ActiveReadServer "Select * from Counters where Cashup_no= " & CashupNo
    End If
    If rs.RecordCount > 0 Then
        If cmdHist.Visible = True Then
            lblClockinTime.Caption = rs.Fields("Shift_Start")
            lblCash(2).Caption = Format(Int(DateDiff("n", lblClockinTime.Caption, rs.Fields("Date_Time")) / 60), "00") & "hrs " & Format(Int(60 * ((DateDiff("n", lblClockinTime, rs.Fields("Date_Time")) / 60) - Int(DateDiff("n", lblClockinTime, rs.Fields("Date_Time")) / 60))), "00") & "min"
        Else
            If IsNull(rs.Fields("Date_Time")) = True Then
                lblCash(2).Caption = Format(Int(DateDiff("n", lblClockinTime.Caption, Now) / 60), "00") & "hrs " & Format(Int(60 * ((DateDiff("n", lblClockinTime, Now) / 60) - Int(DateDiff("n", lblClockinTime, Now) / 60))), "00") & "min"
            Else
                lblCash(2).Caption = Format(Int(DateDiff("n", lblClockinTime.Caption, rs.Fields("Date_Time")) / 60), "00") & "hrs " & Format(Int(60 * ((DateDiff("n", lblClockinTime, rs.Fields("Date_Time")) / 60) - Int(DateDiff("n", lblClockinTime, rs.Fields("Date_Time")) / 60))), "00") & "min"
            End If
        End If
        lblCashupInfo.Caption = "Cashup No: " & rs.Fields("Cashup_No") & " - Clocked In at " & Format(ClockinTime, "HH:mm") & " on " & Format(ClockinTime, "MMM DD YYYY")
        lblCashupInfo.Tag = rs.Fields("Cashup_No")
        If cmdHist.Visible = False Then
            lblCash(51).Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
        Else
            lblCash(51).Caption = Format(rs.Fields("Date_Time"), "dd MMMM yyyy DDD") & " " & Format(rs.Fields("Date_Time"), "HH:MM:SS")
        End If
        ActiveReadServer1 "Select (Select min(Invoice_No) from Sales_Journal where Cashup_No=" & rs.Fields("Cashup_No") & " and User_No=" & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1) & ") as Open_Trans," & _
        "(Select max(Invoice_No) from Sales_Journal where Cashup_No=" & rs.Fields("Cashup_No") & " and User_No=" & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1) & ") as Close_Trans"
        DoEvents
        lblCash(0).Caption = Val(rs1.Fields("Open_Trans") & "")
        lblCash(1).Caption = Val(rs1.Fields("Close_Trans") & "")
        rs1.Close
        lblCash(4).Caption = Val(rs.Fields("Customer_Count") & "")
        lblCash(50).Caption = Val(rs.Fields("No_Sales") & "")
        lblCash(5) = Format(Val(rs.Fields("Cash_Sales_Value") & ""), "0.00")
        lblCash(6) = Format(Val(rs.Fields("Card_Sales_Value") & ""), "0.00")
        lblCash(7) = Format(Val(rs.Fields("Cheque_Sales_Value") & ""), "0.00")
        lblCash(8) = Format(Val(rs.Fields("Charge_Sales_Value") & ""), "0.00")
        lblCash(9) = Format(Val(rs.Fields("Loyalty_Sales_Value") & ""), "0.00")
        lblCash(15).Caption = 0
        For i = 5 To 9
            lblCash(15).Caption = Format(Val(lblCash(15).Caption) + Val(lblCash(i) & ""), "0.00")
        Next i
        
        lblCash(10) = Val(rs.Fields("Cash_Sales_Qty") & "")
        lblCash(11) = Val(rs.Fields("Card_Sales_Qty") & "")
        lblCash(12) = Val(rs.Fields("Cheque_Sales_Qty") & "")
        lblCash(13) = Val(rs.Fields("Charge_Sales_Qty") & "")
        lblCash(14) = Val(rs.Fields("Loyalty_Sales_Qty") & "")
        lblCash(3).Caption = Val(rs.Fields("Cash_Sales_Qty") & "") + Val(rs.Fields("Card_Sales_Qty") & "") + Format(Val(rs.Fields("Cheque_Sales_Qty") & ""), "0.00") + Format(Val(rs.Fields("Charge_Sales_Qty") & ""), "0.00") + Format(Val(rs.Fields("Loyalty_Sales_Qty") & ""), "0.00")
        lblCash(16).Caption = 0
        For i = 10 To 14
            lblCash(16).Caption = Val(lblCash(16).Caption) + Val(lblCash(i) & "")
        Next i
        
        lblCash(28) = Format(Val(rs.Fields("Discount_Perc_Value") & ""), "0.00")
        lblCash(34) = Val(rs.Fields("Discount_Perc_Qty") & "")
        lblCash(39) = Format(Val(rs.Fields("PayOuts_Value") & ""), "0.00")
        lblCash(40) = Format(Val(rs.Fields("Pickups_Value") & ""), "0.00")
        lblCash(41) = Format(Val(rs.Fields("Loans_Value") & ""), "0.00")
        lblCash(42) = Format(Val(rs.Fields("ReceivedonAccount_Value") & ""), "0.00")
        lblCash(48) = Format(Val(rs.Fields("Counted") & ""), "0.00")
        
        lblCash(43) = Val(rs.Fields("PayOuts_Qty") & "")
        lblCash(44) = Val(rs.Fields("Pickups_Qty") & "")
        lblCash(45) = Val(rs.Fields("Loans_Qty") & "")
        lblCash(46) = Val(rs.Fields("ReceivedonAccount_Qty") & "")
        
        lblCash(47) = Format(Val(lblCash(15)) - Val(lblCash(39)) - Val(lblCash(40)) + Val(lblCash(41)) + Val(lblCash(42)), "0.00")
        
        lblCash(49) = Format(Val(lblCash(48)) - Val(lblCash(47)), "0.00")
        
        lblCash(35) = Format(Val(rs.Fields("TaxableSales_Value") & ""), "0.00")
        lblCash(36) = Format(Val(rs.Fields("TotalExcemptSales") & ""), "0.00")
        lblCash(38) = Format(Val(rs.Fields("TotalCalculatedTax_Value") & ""), "0.00")
        lblCash(37) = Format(Val(rs.Fields("TotalCollectedTax_Value") & ""), "0.00")
        lblCash(38) = Format(Val(rs.Fields("TotalCollectedTax_Value") & ""), "0.00")
        lblCash(23) = Format(Val(rs.Fields("Item_Corrects_Value") & ""), "0.00")
        lblCash(29) = Val(rs.Fields("Item_Corrects_Qty") & "")
        lblCash(24) = Format(Val(rs.Fields("Voids_Value") & ""), "0.00")
        lblCash(30) = Val(rs.Fields("Voids_Qty") & "")
        lblCash(25) = Format(Val(rs.Fields("RTMD_Value") & ""), "0.00")
        lblCash(31) = Val(rs.Fields("RTMD_Qty") & "")
        lblCash(26) = Format(Val(rs.Fields("Ullage_Value") & ""), "0.00")
        lblCash(32) = Val(rs.Fields("Ullage_Qty") & "")
        lblCash(19) = Format(Val(rs.Fields("Tipp") & ""), "0.00")
        lblCash(18) = Val(rs.Fields("Tipp_Count") & "")
    End If
    rs.Close
    If cmdHist.Visible = True Then
        cmdHist.Visible = False
        picHist.Visible = False
        On Error GoTo 0
        Exit Sub
    End If
    fmClock.Visible = False
    ActiveReadServer "Select count(Table_No) as TableCount from Table_Listing_View where User_No = " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)
    If rs.Fields("TableCount") > 0 Then
        Select Case rs.Fields("TableCount")
            Case 1
                fmClock.Visible = True
                fmClock.Caption = "This User still has a Open Table."
            Case Else
                fmClock.Visible = True
                fmClock.Caption = "This User still has " & rs.Fields("TableCount") & " Open Tables."
        End Select
        rs.Close
        On Error GoTo 0
        Exit Sub
    End If
    rs.Close
    
    ActiveReadServer "Select count(Tab_No) as TabCount from Tab_Listing_View where User_No = " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)
    If rs.Fields("TabCount") > 0 Then
        Select Case rs.Fields("TabCount")
            Case 1
                fmClock.Visible = True
                fmClock.Caption = "This User still has a Open Tab."
            Case Else
                fmClock.Visible = True
                fmClock.Caption = "This User still has " & rs.Fields("TabCount") & " Open Tabs."
        End Select
        rs.Close
        On Error GoTo 0
        Exit Sub
    End If
    
    rs.Close
    On Error GoTo 0
End Sub
Private Sub Clear_Cashup()
    lblCash(51).Caption = ""
    lblCash(0).Caption = ""
    lblCash(1).Caption = ""
    lblCash(2).Caption = "0 hrs 0 min"
    lblCash(3).Caption = "0"
    lblCash(4).Caption = "0"
    lblCash(50).Caption = "0"
    lblCash(5) = "0.00"
    lblCash(6) = "0.00"
    lblCash(7) = "0"
    lblCash(8) = "0"
    lblCash(9) = "0"
    lblCash(15).Caption = 0
    For i = 5 To 9
        lblCash(15).Caption = "0.00"
    Next i
    lblCash(10) = "0"
    lblCash(11) = "0"
    lblCash(12) = "0"
    lblCash(13) = "0"
    lblCash(14) = "0"
    lblCash(16).Caption = "0"
    lblCash(18).Caption = "0"
    For i = 10 To 14
        lblCash(16).Caption = "0"
    Next i
    lblCash(19) = "0.00"
    lblCash(35) = "0.00"
    lblCash(36) = "0.00"
    lblCash(38) = "0.00"
    lblCash(37) = "0.00"
    
    lblCash(23) = "0.00"
    lblCash(29) = "0"
    lblCash(24) = "0.00"
    lblCash(30) = "0"
    lblCash(25) = "0.00"
    lblCash(31) = "0"
    
    lblCash(39) = "0.00"
    lblCash(40) = "0.00"
    lblCash(41) = "0.00"
    lblCash(42) = "0.00"
    lblCash(48) = "0.00"
    
    lblCash(43) = "0"
    lblCash(44) = "0"
    lblCash(45) = "0"
    lblCash(46) = "0"
    
    lblCash(47) = "0.00"
    
    lblCash(49) = "0.00"
End Sub
Private Sub BtnEnh1_Click(Index As Integer)
    If Timer2.Enabled = True Then Exit Sub
    Select Case grdMain.Visible
         Case True
            grdMain.Visible = False
         Case False
            grdMain.Rows = 0
            grdMain.Visible = True
            ActiveReadServer "Select * from Cash_Ups"
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("User_No") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
                rs.MoveNext
            Wend
            rs.Close
            If grdMain.Rows > 0 Then
                grdMain.Height = 650 * grdMain.Rows - 1
                If grdMain.Rows > 0 Then grdMain.Row = 0
                grdMain.SetFocus
            Else
                lblCashupInfo.Caption = "No Cash'up Available"
                grdMain.Visible = False
            End If
    End Select
End Sub
Private Sub cmdCash_Click(Index As Integer)
    DoEvents
    If cmdCash(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        grdCash.Row = grdCash.Row + 1
        For i = 0 To 15
            If grdCash.TextMatrix(grdCash.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdCash(i).Caption = ""
                    cmdCash(i).TextDescrCB.Text = ""
                    cmdCash(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdCash(i).Visible = False Then cmdCash(i).Visible = True
                Else
                    cmdCash(i).Caption = ""
                    cmdCash(i).TextDescrCB.Text = ""
                    cmdCash(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdCash(i).Visible = False Then cmdCash(i).Visible = True
                    grdCash.Row = grdCash.Row - 1
                    Exit For
                End If
            Else
                cmdCash(i).Caption = "Cashup No: " & grdCash.TextMatrix(grdCash.Row, 0)
                cmdCash(i).Tag = grdCash.TextMatrix(grdCash.Row, 1)
                cmdCash(i).TextDescrCB.Text = grdCash.TextMatrix(grdCash.Row, 2)
            End If
            If grdCash.Row = grdCash.Rows - 1 Then Exit For
            grdCash.Row = grdCash.Row + 1
        Next i
        For b = i + 1 To cmdCash.Count - 1
            cmdCash(b).Caption = "1"
            cmdCash(b).ToolTipText = ""
            cmdCash(i).TextDescrCB.Text = ""
            cmdCash(b).Tag = ""
            cmdCash(b).Visible = False
        Next b
        Exit Sub
    End If
    If cmdCash(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdCash(0).Picture = ""
        While grdCash.TextMatrix(grdCash.Row, 0) <> "ßArrow"
            grdCash.Row = grdCash.Row - 1
        Wend
        grdCash.Row = grdCash.Row - 15
        For i = 0 To 15
            If grdCash.TextMatrix(grdCash.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdCash(i).Caption = ""
                    cmdCash(i).TextDescrCB.Text = ""
                    cmdCash(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdCash(i).Visible = False Then cmdCash(i).Visible = True
                Else
                    cmdCash(i).Caption = ""
                    cmdCash(i).TextDescrCB.Text = ""
                    cmdCash(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdCash(i).Visible = False Then cmdCash(i).Visible = True
                    grdCash.Row = grdCash.Row - 1
                    Exit For
                End If
            Else
                cmdCash(i).Caption = "Cashup No: " & grdCash.TextMatrix(grdCash.Row, 0)
                cmdCash(i).Tag = grdCash.TextMatrix(grdCash.Row, 1)
                cmdCash(i).TextDescrCB.Text = grdCash.TextMatrix(grdCash.Row, 2)
                If cmdCash(i).Visible = False Then cmdCash(i).Visible = True
            End If
            If grdCash.Row = grdCash.Rows - 1 Then Exit For
            grdCash.Row = grdCash.Row + 1
        Next i
        For b = i + 1 To cmdCash.Count - 1
            cmdCash(b).Caption = "1"
            cmdCash(b).Tag = ""
            cmdCash(b).ToolTipText = ""
            cmdCash(b).TextDescrCB.Text = ""
            cmdCash(b).Visible = False
        Next b
        Exit Sub
    End If
    grdMain.Visible = False
    fmClock.Visible = False
    Clear_Cashup
    cmdUser.Caption = cmdCash(Index).TextDescrCB.Text
    LoadCashUp cmdCash(Index).Tag
End Sub

Private Sub cmdErr_Click()
    Timer2.Enabled = False
    cmdErr.Caption = ""
    cmdErr.Visible = False
End Sub

Private Sub cmdFancy_Click(Index As Integer)
    Screen.MousePointer = 11
    If Timer2.Enabled = True Then Screen.MousePointer = 1: Exit Sub
    Select Case cmdFancy(Index).Caption
        Case "Print Count Slip"
            PrintCountSlip
        Case "Capture Cashup"
            If cmdUser.Caption = "<None>" Then
                Timer2.Enabled = True
                cmdErr.Caption = "Select a Cashup First"
                cmdErr.Visible = True
                Screen.MousePointer = 1
                Exit Sub
            End If
            CaptureCashup
        Case "View Cashup Capture"
            CaptureCashup
        Case "Print Cashup"
            If cmdUser.Caption = "<None>" Then
                Timer2.Enabled = True
                cmdErr.Caption = "Select a Cashup First"
                cmdErr.Visible = True
                Screen.MousePointer = 1
                Exit Sub
            End If
            PrintCashup 0
        Case "Finalize Cashup"
            If cmdUser.Caption = "<None>" Then
                Timer2.Enabled = True
                cmdErr.Caption = "Select a Cashup First"
                cmdErr.Visible = True
                Screen.MousePointer = 1
                Exit Sub
            End If
            
            ActiveReadServer "Select count(Table_No) as TableCount" & _
            " from Table_Listing_View where User_No = " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)
            If rs.Fields("TableCount") > 0 Then
                Timer2.Enabled = True
                Select Case rs.Fields("TableCount")
                    Case 1
                        cmdErr.Caption = "This User still has an Open Table."
                    Case Else
                        cmdErr.Caption = "This User still has " & rs.Fields("TableCount") & " Open Tables."
                End Select
                cmdErr.Visible = True
                rs.Close
                Screen.MousePointer = 1
                Exit Sub
            End If
            rs.Close
            
            ActiveReadServer "Select count(Tab_No) as TabCount from Tab_Listing_View where User_No = " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)
            If rs.Fields("TabCount") > 0 Then
                Timer2.Enabled = True
                Select Case rs.Fields("TabCount")
                    Case 1
                        cmdErr.Caption = "This User still has an Open Tab."
                    Case Else
                        cmdErr.Caption = "This User still has " & rs.Fields("TabCount") & " Open Tabs."
                End Select
                cmdErr.Visible = True
                rs.Close
                Screen.MousePointer = 1
                Exit Sub
            End If
            rs.Close
            
            ActiveReadServer "Select count(Table_No) as TableCount from Table_Listing_View where User_No <> " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1) & _
            " and Previous_Owner = " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)
            If rs.Fields("TableCount") > 0 Then
                Timer2.Enabled = True
                Select Case rs.Fields("TableCount")
                    Case 1
                        cmdErr.Caption = "A Transfer by this User has nor been accepted."
                    Case Else
                        cmdErr.Caption = rs.Fields("TableCount") & " Transfers by this User has nor been accepted."
                End Select
                cmdErr.Visible = True
                rs.Close
                Screen.MousePointer = 1
                Exit Sub
            End If
            rs.Close
            FinalizeCashup
            Screen.MousePointer = 1
        
    End Select
    Screen.MousePointer = 1
End Sub
Private Sub FinalizeCashup()
    Load frmQuestion
    frmTillReport.Tag = "Not Now"
    frmQuestion.lblCap = "Are you sure you want finalize this Cashup now?"
    frmQuestion.Show vbModal
    Screen.MousePointer = 11
            ActiveReadServer "Select Date_Time from Counters where Cashup_No = " & frmTillReport.lblCashupInfo.Tag
        rs.Close
     Screen.MousePointer = 1
    
    Select Case QAnswer.Caption
     
        Case "Yes"
            fmClock.Caption = cmdUser.Caption & " Clocked Out at " & Format(Time, "HH:MM")
            fmClock.Visible = True
            ActiveReadServer "Select Date_Time from Counters where Cashup_No = " & frmTillReport.lblCashupInfo.Tag
            If rs.Fields("Date_Time") & "" = "" Then
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1) & ",Getdate(),4," & Workstation_No & ")"
            Else
                ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1) & ",'" & rs.Fields("Date_Time") & "',4," & Workstation_No & ")"
            End If
            rs.Close
            ActiveUpdateServer "Update Users set Drawer_No = 0, Clocked_In = 0 where User_No = " & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)
            ActiveReadServer "Select Date_Time from Counters where Cashup_No= " & frmTillReport.lblCashupInfo.Tag
            If rs.Fields("Date_Time") & "" = "" Then
                rs.Close
                ActiveUpdateServer "Update Counters set Date_Time=Getdate(),Finalized=1,Workstation_No=" & Workstation_No & ",Open_Trans_No='" & lblCash(0).Caption & "',Close_Trans_No='" & lblCash(1).Caption & "',Shift_Start = '" & lblClockinTime.Caption & "' where Cashup_No= " & frmTillReport.lblCashupInfo.Tag
            Else
                rs.Close
                ActiveUpdateServer "Update Counters set Finalized=1,Workstation_No=" & Workstation_No & ",Open_Trans_No='" & lblCash(0).Caption & "',Close_Trans_No='" & lblCash(1).Caption & "',Shift_Start = '" & lblClockinTime.Caption & "' where Cashup_No= " & frmTillReport.lblCashupInfo.Tag
            End If
          
            If Slip_Printer = "<None>" Then GoTo Done
            
            
            PrintCashup 1
Done:
            DoEvents
            Clear_Cashup
            lblCashupInfo.Caption = "Select a Date or a User"
            cmdUser.Caption = "<None>"
    End Select
End Sub
Private Sub CaptureCashup()
    If cmdFancy(0).Caption = "View Cashup Capture" Then
        Load frmCapture
        frmCapture.cmdKey(0).Enabled = False
    Else
        Load frmCapture
        frmCapture.cmdKey(0).Enabled = True
    End If
    frmCapture.Show vbModal
End Sub
Private Sub PrintCashup(Action)
    If QCash = 1 And cmdFancy(2).Enabled = True Then Exit Sub
    Dim x As Printer
    On Error GoTo trap
    frmTillReport.Tag = "Not Now"
    PrintErr = 0
    Slip_Port = ""
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                'Open Slip_Printer For Output As filenum
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                If Slip_Port = "FILE:" Then
                    Open "C:\" & x.DeviceName & ".txt" For Output As filenum
                Else
                    Open Slip_Port For Output As filenum
                End If
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    On Error Resume Next
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(1);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, UCase(Branch_Name)
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Select Case Action
        Case 0
            If cmdFancy(0).Caption = "Capture Cashup" Then
                Print #filenum, "CASHUP" & " X " & "- REPORT"
            Else
                Print #filenum, "CASHUP" & " Z " & "- REPORT"
            End If
        Case 1
            Print #filenum, "CASHUP" & " Z " & "- REPORT"
    End Select
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, UCase(cmdUser.Caption)
    Print #filenum, UCase(lblCash(51))
    Print #filenum, "Cashup No: " & lblCashupInfo.Tag
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);

    Print #filenum, String(33, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, "Cash Sales: " & String(9 - Len(lblCash(5)), " ") & lblCash(5) & "   " & lblCash(10) & String(4 - Len(lblCash(10)), " ")
    
  
    
    Print #filenum, "Card Sales: " & String(9 - Len(lblCash(6)), " ") & lblCash(6) & "   " & lblCash(11) & String(4 - Len(lblCash(11)), " ")
    Print #filenum, "Voucher Sales: " & String(9 - Len(lblCash(7)), " ") & lblCash(7) & "   " & lblCash(12) & String(4 - Len(lblCash(12)), " ")
    Print #filenum, "Charge Sales: " & String(9 - Len(lblCash(8)), " ") & lblCash(8) & "   " & lblCash(13) & String(4 - Len(lblCash(13)), " ")
    Print #filenum, "Loyalty Sales: " & String(9 - Len(lblCash(9)), " ") & lblCash(9) & "   " & lblCash(14) & String(4 - Len(lblCash(14)), " ")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(51) & Chr(18);
    Print #filenum, ""
    Print #filenum, "TOTAL: " & String(9 - Len(lblCash(15)), " ") & lblCash(15) & "   " & lblCash(16) & String(4 - Len(lblCash(16)), " ")
    Print #filenum, ""
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(50);
    Print #filenum, "-Payouts: " & String(9 - Len(lblCash(39)), " ") & lblCash(39) & "   " & lblCash(43) & String(4 - Len(lblCash(43)), " ")
    Print #filenum, "-Pickups: " & String(9 - Len(lblCash(40)), " ") & lblCash(40) & "   " & lblCash(44) & String(4 - Len(lblCash(44)), " ")
    Print #filenum, "+Deposits: " & String(9 - Len(lblCash(41)), " ") & lblCash(41) & "   " & lblCash(45) & String(4 - Len(lblCash(45)), " ")
    Print #filenum, "+Receive on Acc: " & String(9 - Len(lblCash(42)), " ") & lblCash(42) & "   " & lblCash(46) & String(4 - Len(lblCash(46)), " ")
    Print #filenum, String(33, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "TOTAL REPORTED: " & String(9 - Len(lblCash(47)), " ") & lblCash(47)
    Print #filenum, "COUNTED: " & String(9 - Len(lblCash(48)), " ") & lblCash(48)
    Print #filenum, "VARIANCE: " & String(9 - Len(lblCash(49)), " ") & lblCash(49)
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);

    Print #filenum, String(33, "=")
    Print #filenum, "TOTAL IN DRAWER: " & String(9 - Len(lblCash(47)), " ") & lblCash(47)
    cardS = 0
    Charge = 0
    Cheque = 0
    ActiveReadServer "Select * from Counters where Cashup_No = " & lblCashupInfo.Tag
    cardS = Format(Val(rs.Fields("CardsinDrawer_Value") & ""), "0.00")
    If Val(rs.Fields("CardC") & "") <> 0 Then
        cardS = Format(Val(rs.Fields("CardC") & ""), "0.00")
    End If
    Charge = Format(Val(rs.Fields("Chargeindrawer_Value") & ""), "0.00")
    If Val(rs.Fields("ChargeC") & "") <> 0 Then
        Charge = Format(Val(rs.Fields("ChargeC") & ""), "0.00")
    End If
    Cheque = Format(Val(rs.Fields("ChequeinDrawer_Value") & ""), "0.00")
    If Val(rs.Fields("ChequeC") & "") <> 0 Then
        Cheque = Format(Val(rs.Fields("ChequeC") & ""), "0.00")
    End If
    Print #filenum, "-CARD VALUE: " & String(9 - Len(cardS), " ") & cardS
    Print #filenum, "-CHARGE VALUE: " & String(9 - Len(Charge), " ") & Charge
    Print #filenum, "-VOUCHER VALUE: " & String(9 - Len(Cheque), " ") & Cheque
    Print #filenum, String(9, "-")
    Cash = 0
    Cash = Format(Val(lblCash(47)) - cardS - Charge - Cheque, "0.00")
    Print #filenum, "CASH VALUE: " & String(9 - Len(Cash), " ") & Cash
    rs.Close
    deduction = 0
    ActiveReadServer "Select isnull(Comm2,0) as Comm2,Comm1,Com_Calc from Users where User_No=" & Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)
    If rs.RecordCount > 0 Then
        If Val(rs.Fields("Com_Calc") & "") = 0 Then
            Commision = Format(Val(lblCash(47)) * (rs.Fields("Comm1") / 100), "0.00")
        Else
            Commision = Format((Val(lblCash(47)) - lblCash(37)) * (rs.Fields("Comm1") / 100), "0.00")
        End If
        deduction = Val(rs.Fields("Comm2") & "")
    Else
        Commision = "0.00"
        deduction = 0
    End If
    rs.Close
    Print #filenum, "-Commision: " & String(9 - Len(Commision), " ") & Commision
    Due = 0
    Due = Format(Cash - Commision, "0.00")
    Tipp = 0
    If deduction <> 0 Then
        ActiveReadServer1 "Select isnull(Sum(Ave_Cost-Line_Total),0) as  Tipp from Sales_Journal where (Invoice_No > " & Val(lblCash(0)) - 1 & " AND Invoice_No < " & Val(lblCash(1)) + 1 & ") AND (Function_Key = 10) and (User_No = " & Val(Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)) & ")"
        If rs1.RecordCount > 0 Then
            Tipp = rs1.Fields("Tipp") - (rs1.Fields("Tipp") * ((100 - deduction) / 100))
        End If
        rs1.Close
    End If
    If Tipp <> 0 Then
        Print #filenum, "+Service Charge Portion" & String(10, " ")
        Print #filenum, "(Card Sales @ " & deduction & "%): " & String(9 - Len(Format(Tipp, "0.00")), " ") & Format(Tipp, "0.00")
    End If
    Due = Due + Tipp
    Print #filenum, String(9, "-")
    If Due < 0 Then
        Print #filenum, "CASH TO WAITER: " & String(9 - Len(Due), " ") & Format(Abs(Due), "0.00")
    Else
        Print #filenum, "CASH DUE: " & String(9 - Len(Due), " ") & Format(Abs(Due), "0.00")
    End If
    Print #filenum, String(9, "=")
    Print #filenum, String(33, "=")
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "DEPARTMENT BREAKDOWN"
    Print #filenum, "--------------------"
    ActiveReadServer3 "SELECT User_No, LEFT(Sales_Journal.Department_No, PATINDEX('%-%', Sales_Journal.Department_No) - 1) AS Department_No," & _
    " (Select Dept_Name from Departments where Department_No =LEFT(Sales_Journal.Department_No, PATINDEX('%-%', Sales_Journal.Department_No) - 1))" & _
    " as Department_Name, SUM(Line_Total) AS Line_Total" & _
    " From dbo.Sales_Journal" & _
    " WHERE (isnull(Department_No,'')<>'') and (Line_Total <> 0 ) and   (Invoice_No > " & Val(lblCash(0)) - 1 & " AND Invoice_No < " & Val(lblCash(1)) + 1 & ") AND (Function_Key = 7) AND (Extra <> N'Corr')" & _
    " GROUP BY User_No, LEFT(Department_No, PATINDEX('%-%', Department_No) - 1)" & _
    " Having (User_No = " & Val(Mid(cmdUser.Caption, 1, InStr(cmdUser.Caption, "-") - 1)) & ")"
    While Not rs3.EOF
        Print #filenum, rs3.Fields("Department_Name") & ": " & String(9 - Len(Due), " ") & Format(rs3.Fields("Line_Total"), "0.00")
        rs3.MoveNext
    Wend
    rs3.Close
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(64);
'    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(1);
'    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
'    Print #filenum, "PRESENTED BY:"
'    Print #filenum, ""
'    Print #filenum, ""
'    Print #filenum, ""
'    Print #filenum, "ACCEPTED BY:"
'    Print #filenum, ""
'    Print #filenum, ""
'    Print #filenum, ""
'    Print #filenum, "DATED:"
'    Print #filenum, ""
'    Print #filenum, Chr(27) & Chr(50);
'    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(100) & Chr(7);
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(1);
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #1
    On Error GoTo 0
    frmTillReport.Tag = ""
    Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    DoEvents
    frmError.Show vbModal
    On Error GoTo 0
End Sub
Private Sub PrintCountSlip()
    On Error GoTo trap
    frmTillReport.Tag = "Not Now"
    PrintErr = 0
    Slip_Port = ""
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    Print #filenum, Chr(27) & Chr(69) & Chr(1);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, UCase(Branch_Name)
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "CASH UP COUNT"
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    If cmdUser.Caption <> "<None>" Then
        Print #filenum, String(40, "=")
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        Print #filenum, Chr(27) & Chr(97) & Chr(49);
        Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, UCase(cmdUser.Caption)
        Print #filenum, UCase(lblCash(51))
        Print #filenum, Chr(27) & Chr(69) & Chr(48);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        Print #filenum, String(33, "=")
        Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, Chr(27) & Chr(97) & Chr(50);
    End If
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "NOTES"
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(51) & Chr(18);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X R200-00 =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X R100-00 =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X  R50-00 =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X  R20-00 =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X  R10-00 =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "SUB TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, ""
    Print #filenum, "COINS"
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(51) & Chr(18);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X   R5-00 =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X   R2-00 =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X   R1-00 =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X     50c =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X     20c =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X     10c =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)

    Print #filenum, Chr(218) & String(5, Chr(196)) & Chr(191) & "            " & Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & "            " & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(179) & String(5, " ") & Chr(179) & " X      5c =" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(5, Chr(196)) & Chr(217) & "            " & Chr(192) & String(12, Chr(196)) & Chr(217)

    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "SUB TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, String(33, "-")
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CASH TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CARD TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    If QCash = 0 Then
        Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    Else
        Print #filenum, "(Your Slips)  " & Chr(192) & String(15, Chr(196)) & Chr(217)
    End If
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "VOUCHER TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CHARGE TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    
    If QCash = 0 Then
        Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    Else
        Print #filenum, "(Your Slips)  " & Chr(192) & String(15, Chr(196)) & Chr(217)
    End If
    
    If QCash = 1 Then
        Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
        Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
        Print #filenum, "   PAY OUTS:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    End If
    
    If QCash = 0 Then
        Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    Else
        Print #filenum, "(Your Slips)  " & Chr(192) & String(15, Chr(196)) & Chr(217)
    End If
    
    Print #filenum, String(33, "-")
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "COUNT TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "     - FLOAT:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    If QCash = 1 Then
        Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
        Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
        Print #filenum, "= SUB TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
        Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    End If
    If QCash = 1 Then
        Print #filenum, Chr(201) & String(15, Chr(205)) & Chr(187)
        Print #filenum, "" & Chr(186) & String(15, Chr(32)) & Chr(186)
        Print #filenum, "REPORT TOTAL:  " & Chr(186) & String(15 - Len(lblCash(15).Caption), Chr(32)) & Format(0, lblCash(15).Caption) & Chr(186)
        Print #filenum, Chr(200) & String(15, Chr(205)) & Chr(188)
    Else
        Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
        Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
        Print #filenum, "REPORT TOTAL:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
        Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    End If
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "DIFFERENCE:  " & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, String(33, "-")
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, "PRESENTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "ACCEPTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "DATED:"
    Print #filenum, ""
    Print #filenum, Chr(27) & Chr(50);
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(100) & Chr(7);
    Print #filenum, Chr(27) & Chr(64);
    Print #filenum, Chr(27) & Chr(69) & Chr(1);
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #1
    On Error GoTo 0
    frmTillReport.Tag = "   "
    Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim x As Printer
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    DoEvents
    frmError.Show vbModal
    On Error GoTo 0
End Sub

Private Sub cmdHist_Click()
    lblCashdate.Caption = "Reporting on All Cashup's Active Cashup's on: " & Format(Date, "dd mmm yyyy")
    cmdUser.Caption = "<None>"
    lblCashupInfo.Caption = "Select a Date or a User"
    grdMain.Visible = False
    fmClock.Visible = False
    Clear_Cashup
    cmdHist.Visible = False
    picHist.Visible = False
    cmdFancy(2).Enabled = True
    cmdFancy(3).Enabled = True
    cmdFancy(0).Caption = "Capture Cashup"
End Sub
Private Sub cmdKey_Click(Index As Integer)
    If Timer2.Enabled = True Then Exit Sub
    Select Case Index
        Case 5, 6
            SwitchDate Index
        Case Else
            lblCashdate.Caption = "Reporting on All Cashup's Finalized on: " & Format(Dates(Index - 19), "dd mmm yyyy")
            cmdUser.Caption = "<None>"
            lblCashupInfo.Caption = "Select a Date or a User"
            grdMain.Visible = False
            fmClock.Visible = False
            Clear_Cashup
            cmdHist.Visible = True
            picHist.Visible = True
            cmdFancy(2).Enabled = False
            cmdFancy(3).Enabled = False
            cmdFancy(0).Caption = "View Cashup Capture"
            Load_Cash_Ups Dates(Index - 19)
    End Select
End Sub
Private Sub Load_Cash_Ups(Dater As Date)
    Dim Nameduser As String
    grdCash.Rows = 0
    cmdCash(0).Caption = ""
    cmdCash(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Cash_Ups_Fin where (Date_Time > '" & Dater & " 00:00:00' and Date_Time < '" & Dater & " 23:59:59') and Finalized=1 order by Cashup_No "
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdCash.Rows = grdCash.Rows + 1
        If i < 15 And Not rs.EOF Then
            cmdCash(i).Caption = "Cashup No: " & rs.Fields("Cashup_No")
            cmdCash(i).Tag = rs.Fields("Cashup_No")
            If cmdCash(i).Visible = False Then cmdCash(i).Visible = True
            grdCash.Row = grdCash.Rows - 1
            grdCash.TextMatrix(grdCash.Rows - 1, 0) = rs.Fields("Cashup_No")
            grdCash.TextMatrix(grdCash.Rows - 1, 1) = rs.Fields("Cashup_No")
            grdCash.TextMatrix(grdCash.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            cmdCash(i).TextDescrCB.OffsetY = -10
            cmdCash(i).TextDescrCB.ColorNormal = &H800000
            cmdCash(i).TextDescrCB.Text = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
        Else
            If b = 0 Then
                grdCash.TextMatrix(grdCash.Rows - 1, 0) = "ßArrow"
                grdCash.Rows = grdCash.Rows + 1
                If i = 15 Then
                    cmdCash(15).Caption = ""
                    cmdCash(15).Picture = App.Path & "\icons\downArr.bmp"
                    cmdCash(i).TextDescrCB.Text = ""
                    If cmdCash(15).Visible = False Then cmdCash(15).Visible = True
                End If
            End If
            b = b + 1
            grdCash.TextMatrix(grdCash.Rows - 1, 0) = rs.Fields("Cashup_No")
            grdCash.TextMatrix(grdCash.Rows - 1, 1) = rs.Fields("Cashup_No")
           
            Nameduser = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
             grdCash.TextMatrix(grdCash.Rows - 1, 2) = Nameduser
            If b = 14 Then b = 0
        End If
        rs.MoveNext
    Wend
    lblCashupInfo.Caption = "You have " & rs.RecordCount & " Finalized Cashups"
    rs.Close
    For b = i + 1 To cmdCash.Count - 1
       cmdCash(b).Caption = "0"
       cmdCash(b).Visible = False
    Next b
End Sub
Private Sub cmdKey_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        For i = 19 To 25
            If i = Index Then
                cmdKey(Index).BackColor = &H80C0FF
            Else
                cmdKey(i).BackColor = &HA3D3E9
            End If
        Next i
End Sub

Private Sub cmdNext_Click(Index As Integer)
    If Timer2.Enabled = True Then Exit Sub
    Select Case Index
        Case 0
            If GlobalMode = TillMode.CashupMode Then
                frmSplash.Show
            Else
                Select Case UserRecord.uType
                    Case 3
                        Finalizing = False
                        frmSales1.Show
                    Case 4
                        Finalizing = False
                        frmBar.Show
                    Case Else
                        Finalizing = False
                        frmBar.Show
                End Select
            End If
            DoEvents
            Me.Hide
    End Select
End Sub
Private Sub SwitchDate(Action)
      DoEvents
      Select Case Action
            Case 5
                  For i = 25 To 20 Step -1
                        cmdKey(i).Caption = cmdKey(i - 1).Caption
                        cmdKey(i).BackColor = cmdKey(i - 1).BackColor
                        Dates(i - 19) = Dates(i - 20)
                  Next i
                  cmdKey(19).BackColor = &HA3D3E9
                  Dates(0) = DateAdd("d", -1, Dates(1))
                  cmdKey(19).Caption = Format(Dates(0), "DDD")
            Case 6
                  For i = 19 To 24
                        cmdKey(i).Caption = cmdKey(i + 1).Caption
                        cmdKey(i).BackColor = cmdKey(i + 1).BackColor
                        Dates(i - 19) = Dates(i - 18)
                  Next i
                  cmdKey(24).BackColor = &HA3D3E9
                  Dates(6) = DateAdd("d", 1, Dates(5))
                  cmdKey(25).Caption = Format(Dates(6), "DDD")
      End Select
      For i = 0 To 6
            If Dates(i) = Date Then
                 cmdKey(19 + i).Caption = "Today"
                 cmdKey(19 + i).BackColor = &H80C0FF
            End If
      Next i
      If cmdKey(25).Caption <> "Today" Then
          cmdKey(25).BackColor = &HA3D3E9
      End If
      lblTDate = Format(Dates(0), "dd mmm yyyy")
      lblBDate = Format(Dates(6), "dd mmm yyyy")
End Sub

Private Sub cmdUser_Click()
    If Timer2.Enabled = True Then Exit Sub
    Select Case grdMain.Visible
         Case True
            grdMain.Visible = False
         Case False
            grdMain.Rows = 0
            grdMain.Visible = True
            ActiveReadServer "Select * from Cash_Ups"
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("User_No") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
                rs.MoveNext
            Wend
            rs.Close
            If grdMain.Rows > 0 Then
                grdMain.Height = 650 * grdMain.Rows - 1
                If grdMain.Rows > 0 Then grdMain.Row = 0
                grdMain.SetFocus
            Else
                lblCashupInfo.Caption = "No Cash'up Available"
                grdMain.Visible = False
            End If
    End Select
End Sub
Private Sub Form_Activate()
    If frmTillReport.Tag = "Not Now" Then
        frmTillReport.Tag = ""
        Exit Sub
    End If
    If Me.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.79
            Me.Controls(i).Left = Me.Controls(i).Left * 0.785
            Me.Controls(i).Height = Me.Controls(i).Height * 0.784
            Me.Controls(i).top = Me.Controls(i).top * 0.784
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
        On Error GoTo 0
        newBack.Width = Me.Width
        newBack.Height = Me.Height
    End If
    grdMain.Visible = False
    grdMain.ColHidden(1) = True
    lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
    lblCashupInfo.Caption = "Select a Date or a User"
    cmdUser.Caption = "<None>"
    fmClock.Visible = False
    cmdHist.Visible = False
    picHist.Visible = False
    cmdFancy(2).Enabled = True
    cmdFancy(3).Enabled = True
    cmdFancy(0).Caption = "Capture Cashup"
    Clear_Cashup
    On Error Resume Next
    picHoldFocus.SetFocus
    On Error GoTo 0
End Sub
Private Sub Form_Load()
    backcount = -1
      For i = 19 To 25
            If cmdKey(i).Caption = Format(Date, "DDD") Or cmdKey(i).Caption = "Today" Then
                  cmdKey(i).BackColor = &H80C0FF
                  cmdKey(i).Caption = "Today"
                  Dates(i - 19) = Date
                  lblCashdate.Caption = "Reporting on All Cashup's Active Cashup's on: " & Format(Date, "dd mmm yyyy")
                  backcount = i
            Else
                  If backcount <> -1 Then Dates(i - 19) = DateAdd("d", 1, Dates(i - 20))
            End If
      Next i
      If backcount <> -1 Then
            For i = backcount - 19 To 0 Step -1
                  If i <> 6 Then
                        Dates(i) = DateAdd("d", -1, Dates(i + 1))
                  End If
            Next i
      End If
      lblTDate = Format(Dates(0), "dd mmm yyyy")
      lblBDate = Format(Dates(6), "dd mmm yyyy")
      cmdUser.Caption = "<None>"
      Clear_Cashup
End Sub

Private Sub grdMain_Click()
    cmdUser.Caption = grdMain.TextMatrix(grdMain.Row, 0)
    grdMain.Visible = False
    fmClock.Visible = False
    Clear_Cashup
    LoadCashUp 0
    lblCashdate.Caption = "Reporting on All Cashup's Active Cashup's on: " & Format(Date, "dd mmm yyyy")
    cmdFancy(2).Enabled = True
    cmdFancy(3).Enabled = True
    cmdFancy(0).Caption = "Capture Cashup"
End Sub
Private Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            cmdUser.Caption = grdMain.TextMatrix(grdMain.Row, 0)
            grdMain.Visible = False
            LoadCashUp 0
    End Select
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
