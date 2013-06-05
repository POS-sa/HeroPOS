VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmSplit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11430
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00ECDDD7&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplit.frx":0000
   ScaleHeight     =   11430
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer scrolTimer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   780
      Top             =   90
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9175
      Left            =   14400
      ScaleHeight     =   9150
      ScaleWidth      =   675
      TabIndex        =   21
      Top             =   1380
      Visible         =   0   'False
      Width           =   705
      Begin btButtonEx.ButtonEx cmdArrow1 
         Height          =   675
         Index           =   0
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1191
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   8421504
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
      Begin btButtonEx.ButtonEx cmdArrow1 
         Height          =   630
         Index           =   1
         Left            =   0
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   8520
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1111
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   8421504
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
      Begin VB.Shape Shape1 
         BackColor       =   &H00F3EADC&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00DCC29C&
         FillStyle       =   2  'Horizontal Line
         Height          =   8475
         Left            =   15
         Top             =   510
         Width           =   705
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   3060
      Index           =   1
      Left            =   7890
      TabIndex        =   8
      Top             =   1380
      Width           =   3585
      _cx             =   6324
      _cy             =   5397
      Appearance      =   0
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14741752
      BackColorAlternate=   12377839
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   20
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   630
   End
   Begin VB.Timer scrolTimer 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9195
      Left            =   360
      ScaleHeight     =   9165
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   1380
      Width           =   6105
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   675
         Index           =   0
         Left            =   5355
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1191
         Appearance      =   3
         BackColor       =   7848417
         BorderColor     =   2130608
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
      Begin VSFlex8Ctl.VSFlexGrid grdMain 
         Height          =   9150
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   5355
         _cx             =   9446
         _cy             =   16140
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   7453147
         BackColorAlternate=   12377839
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
         Rows            =   20
         Cols            =   20
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   660
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
         WallPaper       =   "frmSplit.frx":BFA0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   630
         Index           =   1
         Left            =   5355
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   8550
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1111
         Appearance      =   3
         BackColor       =   7848417
         BorderColor     =   2130608
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
      Begin VB.Shape picGrid 
         BackColor       =   &H00F3EADC&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00DCC29C&
         FillStyle       =   2  'Horizontal Line
         Height          =   8475
         Left            =   5370
         Top             =   510
         Width           =   705
      End
   End
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   960
      Left            =   14130
      TabIndex        =   1
      Top             =   300
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
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
      ShowFocus       =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   3060
      Index           =   3
      Left            =   7890
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   3585
      _cx             =   6324
      _cy             =   5397
      Appearance      =   0
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14741752
      BackColorAlternate=   12377839
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   20
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   3060
      Index           =   4
      Left            =   11475
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   3615
      _cx             =   6376
      _cy             =   5397
      Appearance      =   0
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14741752
      BackColorAlternate=   12377839
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   20
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   3060
      Index           =   5
      Left            =   7890
      TabIndex        =   7
      Top             =   7500
      Visible         =   0   'False
      Width           =   3585
      _cx             =   6324
      _cy             =   5397
      Appearance      =   0
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14741752
      BackColorAlternate=   12377839
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   20
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   3060
      Index           =   2
      Left            =   11475
      TabIndex        =   9
      Top             =   1380
      Visible         =   0   'False
      Width           =   3615
      _cx             =   6376
      _cy             =   5397
      Appearance      =   0
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14741752
      BackColorAlternate=   12377839
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   20
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   3060
      Index           =   6
      Left            =   11475
      TabIndex        =   10
      Top             =   7500
      Visible         =   0   'False
      Width           =   3615
      _cx             =   6376
      _cy             =   5397
      Appearance      =   0
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14741752
      BackColorAlternate=   12377839
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   20
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1740
      Index           =   1
      Left            =   6570
      TabIndex        =   11
      Top             =   8790
      Width           =   1230
      _Version        =   524298
      _ExtentX        =   2170
      _ExtentY        =   3069
      _StockProps     =   66
      Caption         =   "Done"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1543873
      LogPixels       =   96
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplit.frx":FCD1
      textLT          =   "frmSplit.frx":FD39
      textCT          =   "frmSplit.frx":FD51
      textRT          =   "frmSplit.frx":FD69
      textLM          =   "frmSplit.frx":FD81
      textRM          =   "frmSplit.frx":FD99
      textLB          =   "frmSplit.frx":FDB1
      textCB          =   "frmSplit.frx":FDC9
      textRB          =   "frmSplit.frx":FDE1
      colorBack       =   "frmSplit.frx":FDF9
      colorIntern     =   "frmSplit.frx":FE23
      colorMO         =   "frmSplit.frx":FE4D
      colorFocus      =   "frmSplit.frx":FE77
      colorDisabled   =   "frmSplit.frx":FEA1
      colorPressed    =   "frmSplit.frx":FECB
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdToggle1 
      Height          =   1155
      Index           =   0
      Left            =   6540
      TabIndex        =   16
      Top             =   3570
      Width           =   1260
      _Version        =   524298
      _ExtentX        =   2222
      _ExtentY        =   2037
      _StockProps     =   66
      Caption         =   "5"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   32.25
         Charset         =   2
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1543873
      LogPixels       =   96
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplit.frx":FEF5
      textLT          =   "frmSplit.frx":FF57
      textCT          =   "frmSplit.frx":FF6F
      textRT          =   "frmSplit.frx":FF87
      textLM          =   "frmSplit.frx":FF9F
      textRM          =   "frmSplit.frx":FFB7
      textLB          =   "frmSplit.frx":FFCF
      textCB          =   "frmSplit.frx":FFE7
      textRB          =   "frmSplit.frx":FFFF
      colorBack       =   "frmSplit.frx":10017
      colorIntern     =   "frmSplit.frx":10041
      colorMO         =   "frmSplit.frx":1006B
      colorFocus      =   "frmSplit.frx":10095
      colorDisabled   =   "frmSplit.frx":100BF
      colorPressed    =   "frmSplit.frx":100E9
      Orientation     =   2
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdToggle1 
      Height          =   1155
      Index           =   1
      Left            =   6540
      TabIndex        =   17
      Top             =   5400
      Width           =   1260
      _Version        =   524298
      _ExtentX        =   2222
      _ExtentY        =   2037
      _StockProps     =   66
      Caption         =   "6"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   32.25
         Charset         =   2
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1543873
      LogPixels       =   96
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplit.frx":10113
      textLT          =   "frmSplit.frx":10175
      textCT          =   "frmSplit.frx":1018D
      textRT          =   "frmSplit.frx":101A5
      textLM          =   "frmSplit.frx":101BD
      textRM          =   "frmSplit.frx":101D5
      textLB          =   "frmSplit.frx":101ED
      textCB          =   "frmSplit.frx":10205
      textRB          =   "frmSplit.frx":1021D
      colorBack       =   "frmSplit.frx":10235
      colorIntern     =   "frmSplit.frx":1025F
      colorMO         =   "frmSplit.frx":10289
      colorFocus      =   "frmSplit.frx":102B3
      colorDisabled   =   "frmSplit.frx":102DD
      colorPressed    =   "frmSplit.frx":10307
      Orientation     =   4
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdSplit 
      Height          =   705
      Left            =   6540
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   4710
      Width           =   1260
      _Version        =   524298
      _ExtentX        =   2222
      _ExtentY        =   1244
      _StockProps     =   66
      Caption         =   "Split Once"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1543873
      SmoothEdges     =   2
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplit.frx":10331
      textLT          =   "frmSplit.frx":103A5
      textCT          =   "frmSplit.frx":103BD
      textRT          =   "frmSplit.frx":103D5
      textLM          =   "frmSplit.frx":103ED
      textRM          =   "frmSplit.frx":10405
      textLB          =   "frmSplit.frx":1041D
      textCB          =   "frmSplit.frx":10435
      textRB          =   "frmSplit.frx":1044D
      colorBack       =   "frmSplit.frx":10465
      colorIntern     =   "frmSplit.frx":1048F
      colorMO         =   "frmSplit.frx":104B9
      colorFocus      =   "frmSplit.frx":104E3
      colorDisabled   =   "frmSplit.frx":1050D
      colorPressed    =   "frmSplit.frx":10537
      Orientation     =   2
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1740
      Index           =   0
      Left            =   6570
      TabIndex        =   19
      Top             =   7020
      Visible         =   0   'False
      Width           =   1230
      _Version        =   524298
      _ExtentX        =   2170
      _ExtentY        =   3069
      _StockProps     =   66
      Caption         =   "Print Bills"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1543873
      LogPixels       =   96
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplit.frx":10561
      textLT          =   "frmSplit.frx":105D7
      textCT          =   "frmSplit.frx":105EF
      textRT          =   "frmSplit.frx":10607
      textLM          =   "frmSplit.frx":1061F
      textRM          =   "frmSplit.frx":10637
      textLB          =   "frmSplit.frx":1064F
      textCB          =   "frmSplit.frx":10667
      textRB          =   "frmSplit.frx":1067F
      colorBack       =   "frmSplit.frx":10697
      colorIntern     =   "frmSplit.frx":106C1
      colorMO         =   "frmSplit.frx":106EB
      colorFocus      =   "frmSplit.frx":10715
      colorDisabled   =   "frmSplit.frx":1073F
      colorPressed    =   "frmSplit.frx":10769
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdDeptStrip 
      Height          =   1740
      Index           =   2
      Left            =   6570
      TabIndex        =   20
      Top             =   1380
      Visible         =   0   'False
      Width           =   1230
      _Version        =   524298
      _ExtentX        =   2170
      _ExtentY        =   3069
      _StockProps     =   66
      Caption         =   "Zoom Split"
      Enabled         =   0   'False
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   1543873
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmSplit.frx":10793
      textLT          =   "frmSplit.frx":10807
      textCT          =   "frmSplit.frx":1081F
      textRT          =   "frmSplit.frx":10837
      textLM          =   "frmSplit.frx":1084F
      textRM          =   "frmSplit.frx":10867
      textLB          =   "frmSplit.frx":1087F
      textCB          =   "frmSplit.frx":10897
      textRB          =   "frmSplit.frx":108AF
      colorBack       =   "frmSplit.frx":108C7
      colorIntern     =   "frmSplit.frx":108F1
      colorMO         =   "frmSplit.frx":1091B
      colorFocus      =   "frmSplit.frx":10945
      colorDisabled   =   "frmSplit.frx":1096F
      colorPressed    =   "frmSplit.frx":10999
      HollowFrame     =   -1  'True
   End
   Begin MSForms.Label lblTotal 
      Height          =   675
      Left            =   7380
      TabIndex        =   15
      Top             =   570
      Width           =   6315
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "11139;1191"
      FontName        =   "Arial Narrow"
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblHeading 
      Height          =   675
      Left            =   660
      TabIndex        =   14
      Top             =   570
      Width           =   6795
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "11986;1191"
      FontName        =   "Arial Narrow"
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   540
      TabIndex        =   13
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
      FontWeight      =   700
   End
   Begin MSForms.Label lblUser 
      Height          =   285
      Left            =   10440
      TabIndex        =   12
      Top             =   10860
      Width           =   4335
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7646;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Image newBack 
      Height          =   1815
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "2725;3201"
      Picture         =   "frmSplit.frx":109C3
   End
End
Attribute VB_Name = "frmSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdArrow_Click(Index As Integer)
    Select Case Index
        Case 0
            If grdMain(0).Row <> 1 Then
                grdMain(0).Row = grdMain(0).Row - 1
            End If
        Case 1
            If grdMain(0).Row <> grdMain(0).Rows - 1 Then
                grdMain(0).Row = grdMain(0).Row + 1
            End If
    End Select
    grdMain(0).ShowCell grdMain(0).Row, 0
End Sub
Private Sub cmdArrow_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    scrolTimer.Tag = Index
    scrolTimer.Enabled = True
End Sub

Private Sub cmdArrow_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    scrolTimer.Enabled = False
    scrolTimer.Interval = 700
End Sub
Private Sub cmdArrow1_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    ScrolTimer1.Tag = Index
    ScrolTimer1.Enabled = True
End Sub
Private Sub cmdArrow1_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    ScrolTimer1.Enabled = False
    ScrolTimer1.Interval = 700
End Sub
Private Sub cmdArrow1_Click(Index As Integer)
    For i = 1 To 6
        If grdMain(i).Visible = True Then
            gIndex = i
            Exit For
        End If
    Next i
    Select Case Index
        Case 0
            If grdMain(gIndex).Row <> 1 Then
                grdMain(gIndex).Row = grdMain(gIndex).Row - 1
            End If
        Case 1
            If grdMain(gIndex).Row <> grdMain(gIndex).Rows - 1 Then
                grdMain(gIndex).Row = grdMain(gIndex).Row + 1
            End If
    End Select
    grdMain(gIndex).ShowCell grdMain(gIndex).Row, 0
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    If cmdDeptStrip(2).Caption = "Shrink Split" Then cmdDeptStrip_Click (2)
End Sub
Private Sub cmdDeptStrip_Click(Index As Integer)
    If Process_Running = True Then Exit Sub
    Process_Running = True
    Static gheight
    Static gWidth
    Static gLeft
    Static gTop
    Static gFont
    Static gIndex
    Select Case cmdDeptStrip(Index).Caption
        Case "Done"
            If cmdDeptStrip(2).Caption = "Shrink Split" Then
                grdMain(0).ColHidden(18) = False
                grdMain(gIndex).Height = gheight
                grdMain(gIndex).Width = gWidth
                grdMain(gIndex).Left = gLeft
                grdMain(gIndex).top = gTop
                grdMain(gIndex).FontSize = gFont
                grdMain(gIndex).RowHeightMin = 300
                grdMain(gIndex).Cols = 20
                For i = 1 To Val(cmdSplit.Tag)
                    If grdMain(i).BackColor = &HECDDD7 Then
                        grdMain(i).ColWidth(0) = grdMain(i).Width * 0.15
                        grdMain(i).ColWidth(1) = grdMain(i).Width * 0.6
                        grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
                    Else
                        grdMain(i).Visible = True
                    End If
                Next i
                cmdDeptStrip(Index).Caption = "Zoom Split"
            End If
            If TillData.TableNo <> 0 Then
                ActiveReadServer "Select * from Table_Listing_View where Table_No = " & TillData.TableNo
                If rs.RecordCount > 0 Then
                    User_No = rs.Fields("User_No")
                Else
                    User_No = UserRecord.User_Number
                End If
                rs.Close
                ActiveUpdateServer "Delete from Table_Listing where Table_No= " & TillData.TableNo
                DoEvents
                For i = 0 To 6
                    If grdMain(i).Rows > 1 Then
                        For ib = 1 To grdMain(i).Rows - 1
                            If i = 0 Then
                                ActiveUpdateServer "INSERT INTO [Table_Listing]([Table_No],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value, Table_name)" & _
                                " VALUES('" & Val(TillData.TableNo) & "','" & TillData.Covers & "','" & User_No & "','" & Workstation_No & "','" & grdMain(i).TextMatrix(ib, 0) & "','" & grdMain(i).TextMatrix(ib, 1) & "','" & _
                                grdMain(i).TextMatrix(ib, 2) & "','" & grdMain(i).TextMatrix(ib, 3) & "','" & grdMain(i).TextMatrix(ib, 4) & "','" & grdMain(i).TextMatrix(ib, 5) & "','" & grdMain(i).TextMatrix(ib, 6) & "','" & grdMain(i).TextMatrix(ib, 8) & "','" & grdMain(i).TextMatrix(ib, 9) & "','" & grdMain(i).TextMatrix(ib, 10) & "','" & grdMain(i).TextMatrix(ib, 11) & "','" & grdMain(i).TextMatrix(ib, 12) & "','" & grdMain(i).TextMatrix(ib, 13) & "','" & grdMain(i).TextMatrix(ib, 14) & "','" & grdMain(i).TextMatrix(ib, 7) & "'," & TillData.DocNo & ",0,'" & grdMain(i).ValueMatrix(ib, 17) & "'," & grdMain(i).ValueMatrix(ib, 18) & "," & grdMain(i).ValueMatrix(ib, 19) & ",'" & TillData.Table_Name & "')"
                            Else
                                If TillData.Table_Name = "" Then
                                    New_table_name = ""
                                Else
                                    New_table_name = TillData.Table_Name & "-" & i
                                End If
                                ActiveUpdateServer "INSERT INTO [Table_Listing]([Table_No],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value, Table_name)" & _
                                " VALUES('" & Val(TillData.TableNo & "." & i) & "','0','" & User_No & "','" & Workstation_No & "','" & grdMain(i).TextMatrix(ib, 0) & "','" & grdMain(i).TextMatrix(ib, 1) & "','" & _
                                grdMain(i).TextMatrix(ib, 2) & "','" & grdMain(i).TextMatrix(ib, 3) & "','" & grdMain(i).TextMatrix(ib, 4) & "','" & grdMain(i).TextMatrix(ib, 5) & "','" & grdMain(i).TextMatrix(ib, 6) & "','" & grdMain(i).TextMatrix(ib, 8) & "','" & grdMain(i).TextMatrix(ib, 9) & "','" & grdMain(i).TextMatrix(ib, 10) & "','" & grdMain(i).TextMatrix(ib, 11) & "','" & grdMain(i).TextMatrix(ib, 12) & "','" & grdMain(i).TextMatrix(ib, 13) & "','" & grdMain(i).TextMatrix(ib, 14) & "','" & grdMain(i).TextMatrix(ib, 7) & "'," & TillData.DocNo + i / 10 & ",0,'" & grdMain(i).ValueMatrix(ib, 17) & "',0," & grdMain(i).ValueMatrix(ib, 19) & ",'" & New_table_name & "')"
                            End If
                            DoEvents
                        Next ib
                    End If
                Next i
                Me.Hide
                With frmSales1
                    .lblTable = ""
                    .grdMain.Rows = 1
                    .lblCash.Caption = ""
                    .lblTender.Caption = "0.00"
                    .lblKeyRegister = " Orders Placed for Table No: " & TillData.TableNo & " (Split)"
                    .cmdDept(6).Caption = "No Sale"
                    TillData.DocNo = 0
                    TillData.TableNo = 0
                    TillData.Table_Name = ""
                    TillData.Covers = 0
                    TillData.TotDiscount = 0
                    TillData.TotDiscountVal = 0
                    TillData.TotDiscountCount = 0
                    TillData.TotDiscountValCount = 0
                    GlobalMode = TillMode.FinMode
                    DoEvents
                    Select Case UserRecord.Logged_in
                        Case False
                            frmSales1.KeyPreview = False
                            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
                            frmSplash.Show
                            KeyCode = 0
                            KeyRegister = ""
                            frmSales1.picHoldFocus.Tag = "1"
                            .lblKeyRegister = ""
                            DoEvents
                            frmSales1.Hide
                            Exit Sub
                        Case Else
                            frmInput.Show
                    End Select
                End With
            End If
            If TillData.TabNo <> 0 Then
                ActiveReadServer "Select * from Tab_Listing_View where Tab_No = " & TillData.TabNo
                If rs.RecordCount > 0 Then
                    User_No = rs.Fields("User_No")
                Else
                    User_No = UserRecord.User_Number
                End If
                rs.Close
                ActiveUpdateServer "Delete from Tab_Listing where Tab_No= " & TillData.TabNo
                DoEvents
                For i = 0 To 6
                    If grdMain(i).Rows > 1 Then
                        For ib = 1 To grdMain(i).Rows - 1
                            If i = 0 Then
                                ActiveUpdateServer "INSERT INTO [Tab_Listing]([Tab_No],[Tab_Name],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value)" & _
                                " VALUES('" & Val(TillData.TabNo) & "','" & TillData.TabName & "','" & TillData.Covers & "','" & User_No & "','" & Workstation_No & "','" & grdMain(i).TextMatrix(ib, 0) & "','" & grdMain(i).TextMatrix(ib, 1) & "','" & _
                                grdMain(i).TextMatrix(ib, 2) & "','" & grdMain(i).TextMatrix(ib, 3) & "','" & grdMain(i).TextMatrix(ib, 4) & "','" & grdMain(i).TextMatrix(ib, 5) & "','" & grdMain(i).TextMatrix(ib, 6) & "','" & grdMain(i).TextMatrix(ib, 8) & "','" & grdMain(i).TextMatrix(ib, 9) & "','" & grdMain(i).TextMatrix(ib, 10) & "','" & grdMain(i).TextMatrix(ib, 11) & "','" & grdMain(i).TextMatrix(ib, 12) & "','" & grdMain(i).TextMatrix(ib, 13) & "','" & grdMain(i).TextMatrix(ib, 14) & "','" & grdMain(i).TextMatrix(ib, 7) & "'," & TillData.DocNo & ",0,'" & grdMain(i).ValueMatrix(ib, 17) & "'," & grdMain(i).ValueMatrix(ib, 18) & "," & grdMain(i).ValueMatrix(ib, 19) & ")"
                            Else

                                ActiveUpdateServer "INSERT INTO [Tab_Listing]([Tab_No],[Tab_Name],[Covers], [User_No], [Workstation_No], [Qty],[Short_Desc], [Line_Total], [KeyString], [Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2], [Price_Override], [Printed],[Keyregister],[Doc_No],[Locked],[User_Overide],Discount_Amt,Dicount_Value)" & _
                                " VALUES('" & Val(TillData.TabNo & "." & i) & "','" & TillData.TabName & Str(i) & "','0','" & User_No & "','" & Workstation_No & "','" & grdMain(i).TextMatrix(ib, 0) & "','" & grdMain(i).TextMatrix(ib, 1) & "','" & _
                                grdMain(i).TextMatrix(ib, 2) & "','" & grdMain(i).TextMatrix(ib, 3) & "','" & grdMain(i).TextMatrix(ib, 4) & "','" & grdMain(i).TextMatrix(ib, 5) & "','" & grdMain(i).TextMatrix(ib, 6) & "','" & grdMain(i).TextMatrix(ib, 8) & "','" & grdMain(i).TextMatrix(ib, 9) & "','" & grdMain(i).TextMatrix(ib, 10) & "','" & grdMain(i).TextMatrix(ib, 11) & "','" & grdMain(i).TextMatrix(ib, 12) & "','" & grdMain(i).TextMatrix(ib, 13) & "','" & grdMain(i).TextMatrix(ib, 14) & "','" & grdMain(i).TextMatrix(ib, 7) & "'," & TillData.DocNo & "." & i & ",0,'" & grdMain(i).ValueMatrix(ib, 17) & "',0," & grdMain(i).ValueMatrix(ib, 19) & ")"
                            End If
                            DoEvents
                        Next ib
                    End If
                Next i
                Me.Hide
                With frmBar
                    .lblTab = ""
                    .grdMain.Rows = 1
                    .lblCash.Caption = ""
                    .lblTender.Caption = "0.00"
                    .lblKeyRegister = " Orders Placed for Tab: " & TillData.TabName & " (Split)"
                    .cmdFancy(3).Caption = "Create Tab"
                    TillData.DocNo = 0
                    TillData.TabNo = 0
                    TillData.TabName = ""
                    TillData.Covers = 0
                    GlobalMode = TillMode.FinMode
                    DoEvents
                End With
            End If
        Case "Print Bills"
            If cmdDeptStrip(2).Caption = "Shrink Split" Then
                grdMain(0).ColHidden(18) = False
                grdMain(gIndex).Height = gheight
                grdMain(gIndex).Width = gWidth
                grdMain(gIndex).Left = gLeft
                grdMain(gIndex).top = gTop
                grdMain(gIndex).FontSize = gFont
                grdMain(gIndex).RowHeightMin = 300
                grdMain(gIndex).Cols = 18
                For i = 1 To Val(cmdSplit.Tag)
                    If grdMain(i).BackColor = &HECDDD7 Then
                        grdMain(i).ColWidth(0) = grdMain(i).Width * 0.15
                        grdMain(i).ColWidth(1) = grdMain(i).Width * 0.6
                        grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
                    Else
                        grdMain(i).Visible = True
                    End If
                Next i
                cmdDeptStrip(Index).Caption = "Zoom Split"
            End If
        Case "Zoom Split"
            For i = 1 To 6
                If grdMain(i).BackColor = &HECDDD7 Then
                    picScroll.Visible = True
                    gheight = grdMain(i).Height
                    gWidth = grdMain(i).Width
                    gLeft = grdMain(i).Left
                    gTop = grdMain(i).top
                    gFont = grdMain(i).FontSize
                    gIndex = i
                    grdMain(i).top = 1380
                    grdMain(i).Left = 7740
                    grdMain(i).Width = 6520
                    grdMain(i).Height = 9180
                    grdMain(i).FontSize = 13
                    grdMain(i).ColWidth(0) = grdMain(i).Width * 0.15
                    grdMain(i).ColWidth(1) = grdMain(i).Width * 0.6
                    grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
                    grdMain(i).Cols = 19
                    grdMain(i).ColWidth(0) = grdMain(i).Width * 0.13
                    grdMain(i).ColWidth(1) = grdMain(i).Width * 0.55
                    grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
                    grdMain(i).ColDataType(18) = flexDTBoolean
                    grdMain(gIndex).RowHeightMin = 660
                    grdMain(i).SetFocus
                    If grdMain(i).Rows > 1 Then grdMain(i).Row = 1
                    grdMain(i).HighLight = flexHighlightAlways
                Else
                    grdMain(i).Visible = False
                End If
            Next i
            grdMain(0).ColHidden(18) = True
            
            cmdDeptStrip(Index).Caption = "Shrink Split"
        Case "Shrink Split"
            picScroll.Visible = False
            grdMain(0).ColHidden(18) = False
            grdMain(gIndex).HighLight = flexHighlightWithFocus
            grdMain(gIndex).Height = gheight
            grdMain(gIndex).Width = gWidth
            grdMain(gIndex).Left = gLeft
            grdMain(gIndex).top = gTop
            grdMain(gIndex).FontSize = gFont
            grdMain(gIndex).RowHeightMin = 300
            grdMain(gIndex).Cols = 18
            For i = 1 To Val(cmdSplit.Tag)
                If grdMain(i).BackColor = &HECDDD7 Then
                    grdMain(i).ColWidth(0) = grdMain(i).Width * 0.15
                    grdMain(i).ColWidth(1) = grdMain(i).Width * 0.6
                    grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
                Else
                    grdMain(i).Visible = True
                End If
            Next i
            cmdDeptStrip(Index).Caption = "Zoom Split"
    End Select
    Process_Running = False
End Sub
Private Sub cmdToggle1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Val(cmdSplit.Tag) = 1 Then Exit Sub
            If cmdDeptStrip(2).Caption = "Shrink Split" Then cmdDeptStrip_Click (2)
            For i = 1 To grdMain(Val(cmdSplit.Tag)).Rows - 1
                grdMain(0).Rows = grdMain(0).Rows + 1
                For b = 0 To grdMain(Val(cmdSplit.Tag)).Cols - 1
                   grdMain(0).TextMatrix(grdMain(0).Rows - 1, b) = grdMain(Val(cmdSplit.Tag)).TextMatrix(i, b)
                Next b
            Next i
            grdMain(Val(cmdSplit.Tag)).Rows = 1
            grdMain(Val(cmdSplit.Tag)).BackColor = vbWhite
            cmdSplit.Tag = Val(cmdSplit.Tag) - 1
            grdMain(0).SetFocus
            For i = 1 To grdMain(0).Rows - 1
                Subtotal = Subtotal + grdMain(Index).ValueMatrix(i, 2)
            Next i
            Select Case Panel_no
                Case 1
                    lblHeading.Caption = frmSales1.lblTable
                    lblTotal.Caption = "Subtotal: " & Format(Subtotal, "0.00")
                Case 2
                    lblHeading.Caption = frmBar.lblTab
                    lblTotal.Caption = "Subtotal: " & Format(Subtotal, "0.00")
            End Select
        Case 1
            If Val(cmdSplit.Tag) = 6 Then Exit Sub
            If cmdDeptStrip(2).Caption = "Shrink Split" Then cmdDeptStrip_Click (2)
            cmdSplit.Tag = Val(cmdSplit.Tag) + 1
    End Select
    Select Case Val(cmdSplit.Tag)
        Case 1
            cmdSplit.Caption = "Split Once"
            grdMain(1).Width = 7205
            grdMain(1).Height = 9180
            grdMain(1).FontSize = 13
        Case 2
            cmdSplit.Caption = "Split Twice"
            grdMain(1).Width = 3585
            grdMain(2).Width = 3615
            grdMain(1).Height = 9180
            grdMain(2).Height = 9180
            grdMain(1).FontSize = 11
            grdMain(2).FontSize = 11
        Case 3
            cmdSplit.Caption = "Split Into 3"
            grdMain(1).Height = 4590
            grdMain(3).Height = 4590
            grdMain(3).top = 5970
            grdMain(2).Height = 9180
            grdMain(3).FontSize = 11
        Case 4
            cmdSplit.Caption = "Split Into 4"
            grdMain(2).Height = 4590
            grdMain(4).Height = 4590
            grdMain(4).top = 5970
            grdMain(1).Height = 4590
            grdMain(3).Height = 4590
            grdMain(3).top = 5970
            grdMain(4).FontSize = 11
        Case 5
            cmdSplit.Caption = "Split Into 5"
            grdMain(3).top = 4440
            grdMain(4).top = 4440
            grdMain(5).top = 7500
            grdMain(6).top = 7500
            grdMain(1).Height = 3060
            grdMain(2).Height = 3060
            grdMain(3).Height = 3060
            grdMain(4).Height = 3060
            grdMain(5).Height = 3060
            grdMain(6).Height = 3060
            grdMain(5).FontSize = 11
            grdMain(6).FontSize = 11
        Case 6
            cmdSplit.Caption = "Split Into 6"
    End Select
    For i = 1 To Val(cmdSplit.Tag)
        grdMain(i).ColWidth(0) = grdMain(i).Width * 0.15
        grdMain(i).ColWidth(1) = grdMain(i).Width * 0.6
        grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
        grdMain(i).Visible = True
    Next i
    For b = i To grdMain.Count - 1
        grdMain(i).Visible = False
    Next b
End Sub

Private Sub flashTimer_Timer()
    Select Case grdMain(0).Cell(flexcpBackColor, grdMain(0).Row, 0, grdMain(0).Row, 2)
        Case &HFFFF&
            grdMain(0).Cell(flexcpBackColor, grdMain(0).Row, 0, grdMain(0).Row, 2) = &HFFC0C0
        Case &HFFC0C0
            grdMain(0).Cell(flexcpBackColor, grdMain(0).Row, 0, grdMain(0).Row, 2) = &HFFFF&
    End Select
End Sub

Private Sub Form_Activate()
    If Me.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.782
            Me.Controls(i).Left = Me.Controls(i).Left * 0.778
            Me.Controls(i).Height = Me.Controls(i).Height * 0.78
            Me.Controls(i).top = Me.Controls(i).top * 0.78
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
        On Error GoTo 0
        newBack.Width = Me.Width
        newBack.Height = Me.Height
    End If
    cmdSplit.Caption = "Split Once"
    cmdSplit.Tag = 1
    grdMain(0).ColHidden(19) = True
    grdMain(1).Width = 7205
    grdMain(1).Height = 9180
    grdMain(1).FontSize = 13
    grdMain(1).ColWidth(0) = grdMain(1).Width * 0.15
    grdMain(1).ColWidth(1) = grdMain(1).Width * 0.6
    grdMain(1).ColWidth(2) = grdMain(1).Width * 0.25
    grdMain(0).Rows = 1
    grdMain(1).Rows = 1
    For i = 2 To 6
        grdMain(i).Visible = False
        grdMain(i).Rows = 1
        
    Next i
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
    grdMain(0).Rows = 1
    grdMain(0).SetFocus
    Select Case Panel_no
        Case 1
            With frmSales1
                For i = 1 To .grdMain.Rows - 1
                     If .grdMain.TextMatrix(i, 0) <> "" Then
                        If .grdMain.TextMatrix(i, 8) = "" Or .grdMain.TextMatrix(i, 8) = "Return Item" Then
                            grdMain(0).Rows = grdMain(0).Rows + 1
                            For b = 0 To .grdMain.Cols - 1
                               grdMain(0).TextMatrix(grdMain(0).Rows - 1, b) = .grdMain.TextMatrix(i, b)
                            Next b
                        End If
                    End If
                Next i
                grdMain(0).Row = 1
                grdMain(0).ShowCell 1, 0
            End With
        Case 2
            With frmBar
                For i = 1 To .grdMain.Rows - 1
                     If .grdMain.TextMatrix(i, 0) <> "" Then
                         If .grdMain.TextMatrix(i, 8) = "" Or .grdMain.TextMatrix(i, 8) = "Return Item" Then
                            grdMain(0).Rows = grdMain(0).Rows + 1
                            For b = 0 To .grdMain.Cols - 1
                               grdMain(0).TextMatrix(grdMain(0).Rows - 1, b) = .grdMain.TextMatrix(i, b)
                            Next b
                        End If
                    End If
                Next i
                grdMain(0).Row = 1
                grdMain(0).ShowCell 1, 0
            End With
    End Select
    Select Case Panel_no
        Case 1
            lblHeading.Caption = frmSales1.lblTable
            lblTotal.Caption = "Subtotal: " & frmSales1.lblTender
        Case 2
            lblHeading.Caption = frmBar.lblTab
            lblTotal.Caption = "Subtotal: " & frmBar.lblTender
    End Select
End Sub

Private Sub Form_Load()
    grdMain(i).Rows = 1
    For i = 0 To 6
        If i <> 0 Then
            grdMain(i).TextMatrix(0, 0) = " No " & i
        Else
            grdMain(i).TextMatrix(0, 0) = " No "
        End If
        grdMain(i).TextMatrix(0, 1) = "Description"
        grdMain(i).TextMatrix(0, 2) = "Total "
        If i = 0 Then
            grdMain(i).ColWidth(0) = grdMain(i).Width * 0.13
            grdMain(i).ColWidth(1) = grdMain(i).Width * 0.55
            grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
            grdMain(i).ColWidth(18) = grdMain(i).Width * 0.07
            grdMain(i).ColDataType(18) = flexDTBoolean
        Else
            grdMain(i).ColWidth(0) = grdMain(i).Width * 0.15
            grdMain(i).ColWidth(1) = grdMain(i).Width * 0.6
            grdMain(i).ColWidth(2) = grdMain(i).Width * 0.25
            grdMain(i).Rows = 1
        End If
        grdMain(i).ColAlignment(0) = flexAlignLeftCenter
        grdMain(i).ColAlignment(1) = flexAlignLeftCenter
        grdMain(i).ColAlignment(2) = flexAlignRightCenter
        grdMain(i).ColHidden(3) = True
        grdMain(i).ColHidden(4) = True
        grdMain(i).ColHidden(5) = True
        grdMain(i).ColHidden(6) = True
        grdMain(i).ColHidden(7) = True
        grdMain(i).ColHidden(8) = True
        grdMain(i).ColHidden(9) = True
        grdMain(i).ColHidden(10) = True
        grdMain(i).ColHidden(11) = True
        grdMain(i).ColHidden(12) = True
        grdMain(i).ColHidden(13) = True
        grdMain(i).ColHidden(14) = True
        grdMain(i).ColHidden(15) = True
        grdMain(i).ColHidden(16) = True
        grdMain(i).ColHidden(17) = True
        If i = 0 Then grdMain(i).ColHidden(18) = False
        grdMain(i).Cell(flexcpForeColor, 0, 0, 0, 2) = 0
    Next i
End Sub
Private Sub grdMain_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    If Index = 0 Then
        Select Case Panel_no
            Case 1
                lblHeading.Caption = frmSales1.lblTable
                lblTotal.Caption = "Subtotal: " & frmSales1.lblTender
            Case 2
                lblHeading.Caption = frmBar.lblTab
                lblTotal.Caption = "Subtotal: " & frmBar.lblTender
        End Select
    End If
End Sub

Private Sub grdMain_BeforeMouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    If y < grdMain(Index).Rows * grdMain(Index).RowHeightMin Then
        grdMain(Index).Tag = ""
    Else
        grdMain(Index).Tag = "1"
    End If
End Sub
Private Sub grdMain_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            grdMain(0).BackColor = vbWhite
            If cmdDeptStrip(2).Caption = "Shrink Split" Then
                For i = 1 To 6
                    If grdMain(i).Visible = True Then
                        gIndex = i
                        Exit For
                    End If
                Next i
                For i = 1 To grdMain(gIndex).Rows - 1
                    If grdMain(gIndex).ValueMatrix(i, 18) = True Then
                        grdMain(0).Rows = grdMain(0).Rows + 1
                        For b = 0 To grdMain(gIndex).Cols - 2
                           grdMain(0).TextMatrix(grdMain(0).Rows - 1, b) = grdMain(gIndex).TextMatrix(i, b)
                        Next b
                    End If
                Next i
top1:
                For i = 1 To grdMain(gIndex).Rows - 1
                    If grdMain(gIndex).ValueMatrix(i, 18) = True Then
                        grdMain(gIndex).RemoveItem (i)
                        GoTo top1
                    End If
                Next i
                If grdMain(gIndex).Rows = 1 Then
                    If cmdDeptStrip(2).Caption = "Shrink Split" Then cmdDeptStrip_Click (2)
                End If
            End If
            If grdMain(0).Row = 0 Then Exit Sub
            If cmdDeptStrip(2).Caption = "Zoom Split" Then
                If grdMain(Index).Tag = "" Then
                    Select Case grdMain(0).ValueMatrix(grdMain(0).Row, 18)
                        Case True
                            grdMain(0).TextMatrix(grdMain(0).Row, 18) = 0
                            grdMain(0).Cell(flexcpBackColor, grdMain(0).Row, 0, grdMain(0).Row, 18) = vbWhite
                        Case False
                            grdMain(0).TextMatrix(grdMain(0).Row, 18) = 1
                            grdMain(0).Cell(flexcpBackColor, grdMain(0).Row, 0, grdMain(0).Row, 18) = &HC0FFC0
                    End Select
                End If
            End If
            Subtotal = 0
            For i = 1 To grdMain(0).Rows - 1
                Subtotal = Subtotal + grdMain(Index).ValueMatrix(i, 2)
            Next i
            Select Case Panel_no
                Case 1
                    lblHeading.Caption = frmSales1.lblTable
                    lblTotal.Caption = "Subtotal: " & Format(Subtotal, "0.00")
                Case 2
                    lblHeading.Caption = frmBar.lblTab
                    lblTotal.Caption = "Subtotal: " & Format(Subtotal, "0.00")
            End Select
        Case Else
            cmdDeptStrip(2).Enabled = True
            For i = 1 To 6
               grdMain(i).BackColor = vbWhite
            Next i
            grdMain(Index).BackColor = &HECDDD7
            For i = 1 To grdMain(0).Rows - 1
                If grdMain(0).ValueMatrix(i, 18) = True Then
                    grdMain(Index).Rows = grdMain(Index).Rows + 1
                    For b = 0 To grdMain(0).Cols - 2
                       grdMain(Index).TextMatrix(grdMain(Index).Rows - 1, b) = grdMain(0).TextMatrix(i, b)
                    Next b
                End If
            Next i
top:
            For i = 1 To grdMain(0).Rows - 1
                If grdMain(0).ValueMatrix(i, 18) = True Then
                    grdMain(0).RemoveItem (i)
                    GoTo top
                End If
            Next i
            If cmdDeptStrip(2).Caption = "Shrink Split" Then
                If grdMain(Index).Tag = "" Then
                    Select Case grdMain(Index).ValueMatrix(grdMain(Index).Row, 18)
                        Case True
                            grdMain(Index).TextMatrix(grdMain(Index).Row, 18) = 0
                            grdMain(Index).Cell(flexcpBackColor, grdMain(Index).Row, 0, grdMain(Index).Row, 18) = vbWhite
                        Case False
                            grdMain(Index).TextMatrix(grdMain(Index).Row, 18) = 1
                            grdMain(Index).Cell(flexcpBackColor, grdMain(Index).Row, 0, grdMain(Index).Row, 18) = &HC0FFC0
                    End Select
                End If
            End If
            Subtotal = 0
            For i = 1 To grdMain(Index).Rows - 1
                Subtotal = Subtotal + grdMain(Index).ValueMatrix(i, 2)
            Next i
            Select Case Panel_no
                Case 1
                    lblHeading = " Split " & Index & " of Table No: " & TillData.TableNo
                    lblTotal.Caption = "Subtotal: " & Format(Subtotal, "0.00")
                Case 2
                    lblHeading = " Split " & Index & " of Tab: " & TillData.TabName
                    lblTotal.Caption = "Subtotal: " & Format(Subtotal, "0.00")
            End Select
    End Select
    On Error GoTo 0
End Sub

Private Sub scrolTimer_Timer()
    scrolTimer.Interval = 50
    Select Case scrolTimer.Tag
        Case "0"
            If grdMain(0).Row <> 1 Then
                grdMain(0).Row = grdMain(0).Row - 1
            End If
        Case "1"
            If grdMain(0).Row <> grdMain(0).Rows - 1 Then
                grdMain(0).Row = grdMain(0).Row + 1
            End If
    End Select
    grdMain(0).ShowCell grdMain(0).Row, 0
End Sub

Private Sub ScrolTimer1_Timer()
    For i = 1 To 6
        If grdMain(i).Visible = True Then
            gIndex = i
            Exit For
        End If
    Next i
    ScrolTimer1.Interval = 50
    Select Case ScrolTimer1.Tag
        Case "0"
            If grdMain(gIndex).Row <> 1 Then
                grdMain(gIndex).Row = grdMain(gIndex).Row - 1
            End If
        Case "1"
            If grdMain(gIndex).Row <> grdMain(gIndex).Rows - 1 Then
                grdMain(gIndex).Row = grdMain(gIndex).Row + 1
            End If
    End Select
    grdMain(gIndex).ShowCell grdMain(gIndex).Row, 0
End Sub

Private Sub Timer1_Timer()
    lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
End Sub
