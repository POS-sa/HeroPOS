VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmStockTake1 
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
   Picture         =   "frmStockTake1.frx":0000
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
      Cols            =   2
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
      WallPaper       =   "frmStockTake1.frx":10037
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
      ScaleWidth      =   60
      TabIndex        =   65
      Top             =   360
      Visible         =   0   'False
      Width           =   90
      Begin VSFlex8Ctl.VSFlexGrid grdRem 
         Height          =   8880
         Left            =   90
         TabIndex        =   69
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
         WallPaper       =   "frmStockTake1.frx":13D68
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin btButtonEx.ButtonEx cmdClose 
         Height          =   990
         Left            =   4230
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   68
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
         TabIndex        =   71
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
         TabIndex        =   70
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
   Begin VB.PictureBox picMerge 
      Appearance      =   0  'Flat
      BackColor       =   &H007BCCEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6795
      Left            =   390
      ScaleHeight     =   6795
      ScaleWidth      =   7065
      TabIndex        =   74
      Top             =   3780
      Visible         =   0   'False
      Width           =   7065
      Begin btButtonEx.ButtonEx cmdClose1 
         Height          =   690
         Left            =   4650
         TabIndex        =   81
         Top             =   6090
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1217
         Appearance      =   3
         BackColor       =   6208225
         BorderColor     =   2720171
         Caption         =   "Close"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdView 
         Height          =   690
         Left            =   2325
         TabIndex        =   80
         Top             =   6090
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1217
         Appearance      =   3
         BackColor       =   6208225
         BorderColor     =   2720171
         Caption         =   "View Count"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdPrint 
         Height          =   690
         Left            =   0
         TabIndex        =   79
         Top             =   6090
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1217
         Appearance      =   3
         BackColor       =   6208225
         BorderColor     =   2720171
         Caption         =   "Print"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   660
         Index           =   2
         Left            =   0
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   8220
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1164
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
      Begin btButtonEx.ButtonEx cmdArrow1 
         Height          =   720
         Index           =   0
         Left            =   6330
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1270
         Appearance      =   3
         BackColor       =   2720171
         BorderColor     =   2720171
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
      Begin VSFlex8Ctl.VSFlexGrid grdMerge 
         Height          =   6105
         Left            =   -30
         TabIndex        =   75
         Top             =   0
         Width           =   6345
         _cx             =   11192
         _cy             =   10769
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   670
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
         WallPaper       =   "frmStockTake1.frx":15C0B
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin btButtonEx.ButtonEx cmdArrow1 
         Height          =   720
         Index           =   1
         Left            =   6330
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   5370
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1270
         Appearance      =   3
         BackColor       =   2720171
         BorderColor     =   2720171
         Caption         =   "6"
         CaptionOffsetX  =   1
         CaptionOffsetY  =   1
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
         Height          =   6015
         Left            =   6330
         Top             =   30
         Width           =   705
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
      Height          =   1755
      Left            =   390
      TabIndex        =   0
      Top             =   525
      Width           =   1305
      _Version        =   524298
      _ExtentX        =   2302
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Clear Take"
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
      textCaption     =   "frmStockTake1.frx":17AAE
      textLT          =   "frmStockTake1.frx":17B22
      textCT          =   "frmStockTake1.frx":17B3A
      textRT          =   "frmStockTake1.frx":17B52
      textLM          =   "frmStockTake1.frx":17B6A
      textRM          =   "frmStockTake1.frx":17B82
      textLB          =   "frmStockTake1.frx":17B9A
      textCB          =   "frmStockTake1.frx":17BB2
      textRB          =   "frmStockTake1.frx":17BCA
      colorBack       =   "frmStockTake1.frx":17BE2
      colorIntern     =   "frmStockTake1.frx":17C0C
      colorMO         =   "frmStockTake1.frx":17C36
      colorFocus      =   "frmStockTake1.frx":17C60
      colorDisabled   =   "frmStockTake1.frx":17C8A
      colorPressed    =   "frmStockTake1.frx":17CB4
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
      textCaption     =   "frmStockTake1.frx":17CDE
      textLT          =   "frmStockTake1.frx":17D64
      textCT          =   "frmStockTake1.frx":17D7C
      textRT          =   "frmStockTake1.frx":17D94
      textLM          =   "frmStockTake1.frx":17DAC
      textRM          =   "frmStockTake1.frx":17DC4
      textLB          =   "frmStockTake1.frx":17DDC
      textCB          =   "frmStockTake1.frx":17DF4
      textRB          =   "frmStockTake1.frx":17E0C
      colorBack       =   "frmStockTake1.frx":17E24
      colorIntern     =   "frmStockTake1.frx":17E4E
      colorMO         =   "frmStockTake1.frx":17E78
      colorFocus      =   "frmStockTake1.frx":17EA2
      colorDisabled   =   "frmStockTake1.frx":17ECC
      colorPressed    =   "frmStockTake1.frx":17EF6
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   0
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3795
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":17F20
      textLT          =   "frmStockTake1.frx":17F38
      textCT          =   "frmStockTake1.frx":17F50
      textRT          =   "frmStockTake1.frx":17F68
      textLM          =   "frmStockTake1.frx":17F80
      textRM          =   "frmStockTake1.frx":17F98
      textLB          =   "frmStockTake1.frx":17FB0
      textCB          =   "frmStockTake1.frx":17FC8
      textRB          =   "frmStockTake1.frx":17FE0
      colorBack       =   "frmStockTake1.frx":17FF8
      colorIntern     =   "frmStockTake1.frx":18022
      colorMO         =   "frmStockTake1.frx":1804C
      colorFocus      =   "frmStockTake1.frx":18076
      colorDisabled   =   "frmStockTake1.frx":180A0
      colorPressed    =   "frmStockTake1.frx":180CA
      Style           =   1
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   3
      Left            =   390
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":180F4
      textLT          =   "frmStockTake1.frx":1810C
      textCT          =   "frmStockTake1.frx":18124
      textRT          =   "frmStockTake1.frx":1813C
      textLM          =   "frmStockTake1.frx":18154
      textRM          =   "frmStockTake1.frx":1816C
      textLB          =   "frmStockTake1.frx":18184
      textCB          =   "frmStockTake1.frx":1819C
      textRB          =   "frmStockTake1.frx":181B4
      colorBack       =   "frmStockTake1.frx":181CC
      colorIntern     =   "frmStockTake1.frx":181F6
      colorMO         =   "frmStockTake1.frx":18220
      colorFocus      =   "frmStockTake1.frx":1824A
      colorDisabled   =   "frmStockTake1.frx":18274
      colorPressed    =   "frmStockTake1.frx":1829E
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   5
      Left            =   5100
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":182C8
      textLT          =   "frmStockTake1.frx":182E0
      textCT          =   "frmStockTake1.frx":182F8
      textRT          =   "frmStockTake1.frx":18310
      textLM          =   "frmStockTake1.frx":18328
      textRM          =   "frmStockTake1.frx":18340
      textLB          =   "frmStockTake1.frx":18358
      textCB          =   "frmStockTake1.frx":18370
      textRB          =   "frmStockTake1.frx":18388
      colorBack       =   "frmStockTake1.frx":183A0
      colorIntern     =   "frmStockTake1.frx":183CA
      colorMO         =   "frmStockTake1.frx":183F4
      colorFocus      =   "frmStockTake1.frx":1841E
      colorDisabled   =   "frmStockTake1.frx":18448
      colorPressed    =   "frmStockTake1.frx":18472
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   6
      Left            =   390
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6045
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":1849C
      textLT          =   "frmStockTake1.frx":184B4
      textCT          =   "frmStockTake1.frx":184CC
      textRT          =   "frmStockTake1.frx":184E4
      textLM          =   "frmStockTake1.frx":184FC
      textRM          =   "frmStockTake1.frx":18514
      textLB          =   "frmStockTake1.frx":1852C
      textCB          =   "frmStockTake1.frx":18544
      textRB          =   "frmStockTake1.frx":1855C
      colorBack       =   "frmStockTake1.frx":18574
      colorIntern     =   "frmStockTake1.frx":1859E
      colorMO         =   "frmStockTake1.frx":185C8
      colorFocus      =   "frmStockTake1.frx":185F2
      colorDisabled   =   "frmStockTake1.frx":1861C
      colorPressed    =   "frmStockTake1.frx":18646
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   8
      Left            =   5100
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6045
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":18670
      textLT          =   "frmStockTake1.frx":18688
      textCT          =   "frmStockTake1.frx":186A0
      textRT          =   "frmStockTake1.frx":186B8
      textLM          =   "frmStockTake1.frx":186D0
      textRM          =   "frmStockTake1.frx":186E8
      textLB          =   "frmStockTake1.frx":18700
      textCB          =   "frmStockTake1.frx":18718
      textRB          =   "frmStockTake1.frx":18730
      colorBack       =   "frmStockTake1.frx":18748
      colorIntern     =   "frmStockTake1.frx":18772
      colorMO         =   "frmStockTake1.frx":1879C
      colorFocus      =   "frmStockTake1.frx":187C6
      colorDisabled   =   "frmStockTake1.frx":187F0
      colorPressed    =   "frmStockTake1.frx":1881A
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   9
      Left            =   390
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7170
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":18844
      textLT          =   "frmStockTake1.frx":1885C
      textCT          =   "frmStockTake1.frx":18874
      textRT          =   "frmStockTake1.frx":1888C
      textLM          =   "frmStockTake1.frx":188A4
      textRM          =   "frmStockTake1.frx":188BC
      textLB          =   "frmStockTake1.frx":188D4
      textCB          =   "frmStockTake1.frx":188EC
      textRB          =   "frmStockTake1.frx":18904
      colorBack       =   "frmStockTake1.frx":1891C
      colorIntern     =   "frmStockTake1.frx":18946
      colorMO         =   "frmStockTake1.frx":18970
      colorFocus      =   "frmStockTake1.frx":1899A
      colorDisabled   =   "frmStockTake1.frx":189C4
      colorPressed    =   "frmStockTake1.frx":189EE
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   11
      Left            =   5100
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7170
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":18A18
      textLT          =   "frmStockTake1.frx":18A30
      textCT          =   "frmStockTake1.frx":18A48
      textRT          =   "frmStockTake1.frx":18A60
      textLM          =   "frmStockTake1.frx":18A78
      textRM          =   "frmStockTake1.frx":18A90
      textLB          =   "frmStockTake1.frx":18AA8
      textCB          =   "frmStockTake1.frx":18AC0
      textRB          =   "frmStockTake1.frx":18AD8
      colorBack       =   "frmStockTake1.frx":18AF0
      colorIntern     =   "frmStockTake1.frx":18B1A
      colorMO         =   "frmStockTake1.frx":18B44
      colorFocus      =   "frmStockTake1.frx":18B6E
      colorDisabled   =   "frmStockTake1.frx":18B98
      colorPressed    =   "frmStockTake1.frx":18BC2
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   12
      Left            =   390
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8295
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":18BEC
      textLT          =   "frmStockTake1.frx":18C04
      textCT          =   "frmStockTake1.frx":18C1C
      textRT          =   "frmStockTake1.frx":18C34
      textLM          =   "frmStockTake1.frx":18C4C
      textRM          =   "frmStockTake1.frx":18C64
      textLB          =   "frmStockTake1.frx":18C7C
      textCB          =   "frmStockTake1.frx":18C94
      textRB          =   "frmStockTake1.frx":18CAC
      colorBack       =   "frmStockTake1.frx":18CC4
      colorIntern     =   "frmStockTake1.frx":18CEE
      colorMO         =   "frmStockTake1.frx":18D18
      colorFocus      =   "frmStockTake1.frx":18D42
      colorDisabled   =   "frmStockTake1.frx":18D6C
      colorPressed    =   "frmStockTake1.frx":18D96
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   14
      Left            =   5100
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8295
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":18DC0
      textLT          =   "frmStockTake1.frx":18DD8
      textCT          =   "frmStockTake1.frx":18DF0
      textRT          =   "frmStockTake1.frx":18E08
      textLM          =   "frmStockTake1.frx":18E20
      textRM          =   "frmStockTake1.frx":18E38
      textLB          =   "frmStockTake1.frx":18E50
      textCB          =   "frmStockTake1.frx":18E68
      textRB          =   "frmStockTake1.frx":18E80
      colorBack       =   "frmStockTake1.frx":18E98
      colorIntern     =   "frmStockTake1.frx":18EC2
      colorMO         =   "frmStockTake1.frx":18EEC
      colorFocus      =   "frmStockTake1.frx":18F16
      colorDisabled   =   "frmStockTake1.frx":18F40
      colorPressed    =   "frmStockTake1.frx":18F6A
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   15
      Left            =   390
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9420
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":18F94
      textLT          =   "frmStockTake1.frx":18FAC
      textCT          =   "frmStockTake1.frx":18FC4
      textRT          =   "frmStockTake1.frx":18FDC
      textLM          =   "frmStockTake1.frx":18FF4
      textRM          =   "frmStockTake1.frx":1900C
      textLB          =   "frmStockTake1.frx":19024
      textCB          =   "frmStockTake1.frx":1903C
      textRB          =   "frmStockTake1.frx":19054
      colorBack       =   "frmStockTake1.frx":1906C
      colorIntern     =   "frmStockTake1.frx":19096
      colorMO         =   "frmStockTake1.frx":190C0
      colorFocus      =   "frmStockTake1.frx":190EA
      colorDisabled   =   "frmStockTake1.frx":19114
      colorPressed    =   "frmStockTake1.frx":1913E
      Style           =   1
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   17
      Left            =   5100
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9420
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":19168
      textLT          =   "frmStockTake1.frx":19180
      textCT          =   "frmStockTake1.frx":19198
      textRT          =   "frmStockTake1.frx":191B0
      textLM          =   "frmStockTake1.frx":191C8
      textRM          =   "frmStockTake1.frx":191E0
      textLB          =   "frmStockTake1.frx":191F8
      textCB          =   "frmStockTake1.frx":19210
      textRB          =   "frmStockTake1.frx":19228
      colorBack       =   "frmStockTake1.frx":19240
      colorIntern     =   "frmStockTake1.frx":1926A
      colorMO         =   "frmStockTake1.frx":19294
      colorFocus      =   "frmStockTake1.frx":192BE
      colorDisabled   =   "frmStockTake1.frx":192E8
      colorPressed    =   "frmStockTake1.frx":19312
      Style           =   1
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   1
      Left            =   2745
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3795
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":1933C
      textLT          =   "frmStockTake1.frx":19354
      textCT          =   "frmStockTake1.frx":1936C
      textRT          =   "frmStockTake1.frx":19384
      textLM          =   "frmStockTake1.frx":1939C
      textRM          =   "frmStockTake1.frx":193B4
      textLB          =   "frmStockTake1.frx":193CC
      textCB          =   "frmStockTake1.frx":193E4
      textRB          =   "frmStockTake1.frx":193FC
      colorBack       =   "frmStockTake1.frx":19414
      colorIntern     =   "frmStockTake1.frx":1943E
      colorMO         =   "frmStockTake1.frx":19468
      colorFocus      =   "frmStockTake1.frx":19492
      colorDisabled   =   "frmStockTake1.frx":194BC
      colorPressed    =   "frmStockTake1.frx":194E6
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   4
      Left            =   2745
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":19510
      textLT          =   "frmStockTake1.frx":19528
      textCT          =   "frmStockTake1.frx":19540
      textRT          =   "frmStockTake1.frx":19558
      textLM          =   "frmStockTake1.frx":19570
      textRM          =   "frmStockTake1.frx":19588
      textLB          =   "frmStockTake1.frx":195A0
      textCB          =   "frmStockTake1.frx":195B8
      textRB          =   "frmStockTake1.frx":195D0
      colorBack       =   "frmStockTake1.frx":195E8
      colorIntern     =   "frmStockTake1.frx":19612
      colorMO         =   "frmStockTake1.frx":1963C
      colorFocus      =   "frmStockTake1.frx":19666
      colorDisabled   =   "frmStockTake1.frx":19690
      colorPressed    =   "frmStockTake1.frx":196BA
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   7
      Left            =   2745
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6045
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":196E4
      textLT          =   "frmStockTake1.frx":196FC
      textCT          =   "frmStockTake1.frx":19714
      textRT          =   "frmStockTake1.frx":1972C
      textLM          =   "frmStockTake1.frx":19744
      textRM          =   "frmStockTake1.frx":1975C
      textLB          =   "frmStockTake1.frx":19774
      textCB          =   "frmStockTake1.frx":1978C
      textRB          =   "frmStockTake1.frx":197A4
      colorBack       =   "frmStockTake1.frx":197BC
      colorIntern     =   "frmStockTake1.frx":197E6
      colorMO         =   "frmStockTake1.frx":19810
      colorFocus      =   "frmStockTake1.frx":1983A
      colorDisabled   =   "frmStockTake1.frx":19864
      colorPressed    =   "frmStockTake1.frx":1988E
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   10
      Left            =   2745
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7170
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":198B8
      textLT          =   "frmStockTake1.frx":198D0
      textCT          =   "frmStockTake1.frx":198E8
      textRT          =   "frmStockTake1.frx":19900
      textLM          =   "frmStockTake1.frx":19918
      textRM          =   "frmStockTake1.frx":19930
      textLB          =   "frmStockTake1.frx":19948
      textCB          =   "frmStockTake1.frx":19960
      textRB          =   "frmStockTake1.frx":19978
      colorBack       =   "frmStockTake1.frx":19990
      colorIntern     =   "frmStockTake1.frx":199BA
      colorMO         =   "frmStockTake1.frx":199E4
      colorFocus      =   "frmStockTake1.frx":19A0E
      colorDisabled   =   "frmStockTake1.frx":19A38
      colorPressed    =   "frmStockTake1.frx":19A62
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   13
      Left            =   2745
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8295
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":19A8C
      textLT          =   "frmStockTake1.frx":19AA4
      textCT          =   "frmStockTake1.frx":19ABC
      textRT          =   "frmStockTake1.frx":19AD4
      textLM          =   "frmStockTake1.frx":19AEC
      textRM          =   "frmStockTake1.frx":19B04
      textLB          =   "frmStockTake1.frx":19B1C
      textCB          =   "frmStockTake1.frx":19B34
      textRB          =   "frmStockTake1.frx":19B4C
      colorBack       =   "frmStockTake1.frx":19B64
      colorIntern     =   "frmStockTake1.frx":19B8E
      colorMO         =   "frmStockTake1.frx":19BB8
      colorFocus      =   "frmStockTake1.frx":19BE2
      colorDisabled   =   "frmStockTake1.frx":19C0C
      colorPressed    =   "frmStockTake1.frx":19C36
      Style           =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   16
      Left            =   2745
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9420
      Visible         =   0   'False
      Width           =   2355
      _Version        =   524298
      _ExtentX        =   4154
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":19C60
      textLT          =   "frmStockTake1.frx":19C78
      textCT          =   "frmStockTake1.frx":19C90
      textRT          =   "frmStockTake1.frx":19CA8
      textLM          =   "frmStockTake1.frx":19CC0
      textRM          =   "frmStockTake1.frx":19CD8
      textLB          =   "frmStockTake1.frx":19CF0
      textCB          =   "frmStockTake1.frx":19D08
      textRB          =   "frmStockTake1.frx":19D20
      colorBack       =   "frmStockTake1.frx":19D38
      colorIntern     =   "frmStockTake1.frx":19D62
      colorMO         =   "frmStockTake1.frx":19D8C
      colorFocus      =   "frmStockTake1.frx":19DB6
      colorDisabled   =   "frmStockTake1.frx":19DE0
      colorPressed    =   "frmStockTake1.frx":19E0A
      Style           =   1
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
      WallPaper       =   "frmStockTake1.frx":19E34
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   3
      Left            =   2850
      TabIndex        =   20
      Top             =   525
      Width           =   1635
      _Version        =   524298
      _ExtentX        =   2884
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Merge Take"
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
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   10736617
      SpecialEffect   =   1
      CaptionWordWrapPerc=   80
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmStockTake1.frx":1BCD7
      textLT          =   "frmStockTake1.frx":1BD4B
      textCT          =   "frmStockTake1.frx":1BD63
      textRT          =   "frmStockTake1.frx":1BD7B
      textLM          =   "frmStockTake1.frx":1BD93
      textRM          =   "frmStockTake1.frx":1BDAB
      textLB          =   "frmStockTake1.frx":1BDC3
      textCB          =   "frmStockTake1.frx":1BDDB
      textRB          =   "frmStockTake1.frx":1BDF3
      colorBack       =   "frmStockTake1.frx":1BE0B
      colorIntern     =   "frmStockTake1.frx":1BE35
      colorMO         =   "frmStockTake1.frx":1BE5F
      colorFocus      =   "frmStockTake1.frx":1BE89
      colorDisabled   =   "frmStockTake1.frx":1BEB3
      colorPressed    =   "frmStockTake1.frx":1BEDD
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdPlu 
      Height          =   1125
      Index           =   2
      Left            =   5100
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3795
      Visible         =   0   'False
      Width           =   2340
      _Version        =   524298
      _ExtentX        =   4128
      _ExtentY        =   1984
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
      textCaption     =   "frmStockTake1.frx":1BF07
      textLT          =   "frmStockTake1.frx":1BF1F
      textCT          =   "frmStockTake1.frx":1BF37
      textRT          =   "frmStockTake1.frx":1BF4F
      textLM          =   "frmStockTake1.frx":1BF67
      textRM          =   "frmStockTake1.frx":1BF7F
      textLB          =   "frmStockTake1.frx":1BF97
      textCB          =   "frmStockTake1.frx":1BFAF
      textRB          =   "frmStockTake1.frx":1BFC7
      colorBack       =   "frmStockTake1.frx":1BFDF
      colorIntern     =   "frmStockTake1.frx":1C009
      colorMO         =   "frmStockTake1.frx":1C033
      colorFocus      =   "frmStockTake1.frx":1C05D
      colorDisabled   =   "frmStockTake1.frx":1C087
      colorPressed    =   "frmStockTake1.frx":1C0B1
      Style           =   1
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   855
      Index           =   6
      Left            =   4500
      TabIndex        =   22
      Top             =   1425
      Width           =   2925
      _Version        =   524298
      _ExtentX        =   5159
      _ExtentY        =   1508
      _StockProps     =   66
      Caption         =   "Print Count Sheet"
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
      textCaption     =   "frmStockTake1.frx":1C0DB
      textLT          =   "frmStockTake1.frx":1C15D
      textCT          =   "frmStockTake1.frx":1C175
      textRT          =   "frmStockTake1.frx":1C18D
      textLM          =   "frmStockTake1.frx":1C1A5
      textRM          =   "frmStockTake1.frx":1C1BD
      textLB          =   "frmStockTake1.frx":1C1D5
      textCB          =   "frmStockTake1.frx":1C1ED
      textRB          =   "frmStockTake1.frx":1C205
      colorBack       =   "frmStockTake1.frx":1C21D
      colorIntern     =   "frmStockTake1.frx":1C247
      colorMO         =   "frmStockTake1.frx":1C271
      colorFocus      =   "frmStockTake1.frx":1C29B
      colorDisabled   =   "frmStockTake1.frx":1C2C5
      colorPressed    =   "frmStockTake1.frx":1C2EF
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
      textCaption     =   "frmStockTake1.frx":1C319
      textLT          =   "frmStockTake1.frx":1C37B
      textCT          =   "frmStockTake1.frx":1C393
      textRT          =   "frmStockTake1.frx":1C3AB
      textLM          =   "frmStockTake1.frx":1C3C3
      textRM          =   "frmStockTake1.frx":1C3DB
      textLB          =   "frmStockTake1.frx":1C3F3
      textCB          =   "frmStockTake1.frx":1C40B
      textRB          =   "frmStockTake1.frx":1C423
      colorBack       =   "frmStockTake1.frx":1C43B
      colorIntern     =   "frmStockTake1.frx":1C465
      colorMO         =   "frmStockTake1.frx":1C48F
      colorFocus      =   "frmStockTake1.frx":1C4B9
      colorDisabled   =   "frmStockTake1.frx":1C4E3
      colorPressed    =   "frmStockTake1.frx":1C50D
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
      textCaption     =   "frmStockTake1.frx":1C537
      textLT          =   "frmStockTake1.frx":1C599
      textCT          =   "frmStockTake1.frx":1C5B1
      textRT          =   "frmStockTake1.frx":1C5C9
      textLM          =   "frmStockTake1.frx":1C5E1
      textRM          =   "frmStockTake1.frx":1C5F9
      textLB          =   "frmStockTake1.frx":1C611
      textCB          =   "frmStockTake1.frx":1C629
      textRB          =   "frmStockTake1.frx":1C641
      colorBack       =   "frmStockTake1.frx":1C659
      colorIntern     =   "frmStockTake1.frx":1C683
      colorMO         =   "frmStockTake1.frx":1C6AD
      colorFocus      =   "frmStockTake1.frx":1C6D7
      colorDisabled   =   "frmStockTake1.frx":1C701
      colorPressed    =   "frmStockTake1.frx":1C72B
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
      textCaption     =   "frmStockTake1.frx":1C755
      textLT          =   "frmStockTake1.frx":1C7B7
      textCT          =   "frmStockTake1.frx":1C7CF
      textRT          =   "frmStockTake1.frx":1C7E7
      textLM          =   "frmStockTake1.frx":1C7FF
      textRM          =   "frmStockTake1.frx":1C817
      textLB          =   "frmStockTake1.frx":1C82F
      textCB          =   "frmStockTake1.frx":1C847
      textRB          =   "frmStockTake1.frx":1C85F
      colorBack       =   "frmStockTake1.frx":1C877
      colorIntern     =   "frmStockTake1.frx":1C8A1
      colorMO         =   "frmStockTake1.frx":1C8CB
      colorFocus      =   "frmStockTake1.frx":1C8F5
      colorDisabled   =   "frmStockTake1.frx":1C91F
      colorPressed    =   "frmStockTake1.frx":1C949
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
      textCaption     =   "frmStockTake1.frx":1C973
      textLT          =   "frmStockTake1.frx":1C9D5
      textCT          =   "frmStockTake1.frx":1C9ED
      textRT          =   "frmStockTake1.frx":1CA05
      textLM          =   "frmStockTake1.frx":1CA1D
      textRM          =   "frmStockTake1.frx":1CA35
      textLB          =   "frmStockTake1.frx":1CA4D
      textCB          =   "frmStockTake1.frx":1CA65
      textRB          =   "frmStockTake1.frx":1CA7D
      colorBack       =   "frmStockTake1.frx":1CA95
      colorIntern     =   "frmStockTake1.frx":1CABF
      colorMO         =   "frmStockTake1.frx":1CAE9
      colorFocus      =   "frmStockTake1.frx":1CB13
      colorDisabled   =   "frmStockTake1.frx":1CB3D
      colorPressed    =   "frmStockTake1.frx":1CB67
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
      textCaption     =   "frmStockTake1.frx":1CB91
      textLT          =   "frmStockTake1.frx":1CBF3
      textCT          =   "frmStockTake1.frx":1CC0B
      textRT          =   "frmStockTake1.frx":1CC23
      textLM          =   "frmStockTake1.frx":1CC3B
      textRM          =   "frmStockTake1.frx":1CC53
      textLB          =   "frmStockTake1.frx":1CC6B
      textCB          =   "frmStockTake1.frx":1CC83
      textRB          =   "frmStockTake1.frx":1CC9B
      colorBack       =   "frmStockTake1.frx":1CCB3
      colorIntern     =   "frmStockTake1.frx":1CCDD
      colorMO         =   "frmStockTake1.frx":1CD07
      colorFocus      =   "frmStockTake1.frx":1CD31
      colorDisabled   =   "frmStockTake1.frx":1CD5B
      colorPressed    =   "frmStockTake1.frx":1CD85
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
      textCaption     =   "frmStockTake1.frx":1CDAF
      textLT          =   "frmStockTake1.frx":1CE11
      textCT          =   "frmStockTake1.frx":1CE29
      textRT          =   "frmStockTake1.frx":1CE41
      textLM          =   "frmStockTake1.frx":1CE59
      textRM          =   "frmStockTake1.frx":1CE71
      textLB          =   "frmStockTake1.frx":1CE89
      textCB          =   "frmStockTake1.frx":1CEA1
      textRB          =   "frmStockTake1.frx":1CEB9
      colorBack       =   "frmStockTake1.frx":1CED1
      colorIntern     =   "frmStockTake1.frx":1CEFB
      colorMO         =   "frmStockTake1.frx":1CF25
      colorFocus      =   "frmStockTake1.frx":1CF4F
      colorDisabled   =   "frmStockTake1.frx":1CF79
      colorPressed    =   "frmStockTake1.frx":1CFA3
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
      textCaption     =   "frmStockTake1.frx":1CFCD
      textLT          =   "frmStockTake1.frx":1D02F
      textCT          =   "frmStockTake1.frx":1D047
      textRT          =   "frmStockTake1.frx":1D05F
      textLM          =   "frmStockTake1.frx":1D077
      textRM          =   "frmStockTake1.frx":1D08F
      textLB          =   "frmStockTake1.frx":1D0A7
      textCB          =   "frmStockTake1.frx":1D0BF
      textRB          =   "frmStockTake1.frx":1D0D7
      colorBack       =   "frmStockTake1.frx":1D0EF
      colorIntern     =   "frmStockTake1.frx":1D119
      colorMO         =   "frmStockTake1.frx":1D143
      colorFocus      =   "frmStockTake1.frx":1D16D
      colorDisabled   =   "frmStockTake1.frx":1D197
      colorPressed    =   "frmStockTake1.frx":1D1C1
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
      textCaption     =   "frmStockTake1.frx":1D1EB
      textLT          =   "frmStockTake1.frx":1D24D
      textCT          =   "frmStockTake1.frx":1D265
      textRT          =   "frmStockTake1.frx":1D27D
      textLM          =   "frmStockTake1.frx":1D295
      textRM          =   "frmStockTake1.frx":1D2AD
      textLB          =   "frmStockTake1.frx":1D2C5
      textCB          =   "frmStockTake1.frx":1D2DD
      textRB          =   "frmStockTake1.frx":1D2F5
      colorBack       =   "frmStockTake1.frx":1D30D
      colorIntern     =   "frmStockTake1.frx":1D337
      colorMO         =   "frmStockTake1.frx":1D361
      colorFocus      =   "frmStockTake1.frx":1D38B
      colorDisabled   =   "frmStockTake1.frx":1D3B5
      colorPressed    =   "frmStockTake1.frx":1D3DF
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
      textCaption     =   "frmStockTake1.frx":1D409
      textLT          =   "frmStockTake1.frx":1D46B
      textCT          =   "frmStockTake1.frx":1D483
      textRT          =   "frmStockTake1.frx":1D49B
      textLM          =   "frmStockTake1.frx":1D4B3
      textRM          =   "frmStockTake1.frx":1D4CB
      textLB          =   "frmStockTake1.frx":1D4E3
      textCB          =   "frmStockTake1.frx":1D4FB
      textRB          =   "frmStockTake1.frx":1D513
      colorBack       =   "frmStockTake1.frx":1D52B
      colorIntern     =   "frmStockTake1.frx":1D555
      colorMO         =   "frmStockTake1.frx":1D57F
      colorFocus      =   "frmStockTake1.frx":1D5A9
      colorDisabled   =   "frmStockTake1.frx":1D5D3
      colorPressed    =   "frmStockTake1.frx":1D5FD
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
      textCaption     =   "frmStockTake1.frx":1D627
      textLT          =   "frmStockTake1.frx":1D689
      textCT          =   "frmStockTake1.frx":1D6A1
      textRT          =   "frmStockTake1.frx":1D6B9
      textLM          =   "frmStockTake1.frx":1D6D1
      textRM          =   "frmStockTake1.frx":1D6E9
      textLB          =   "frmStockTake1.frx":1D701
      textCB          =   "frmStockTake1.frx":1D719
      textRB          =   "frmStockTake1.frx":1D731
      colorBack       =   "frmStockTake1.frx":1D749
      colorIntern     =   "frmStockTake1.frx":1D773
      colorMO         =   "frmStockTake1.frx":1D79D
      colorFocus      =   "frmStockTake1.frx":1D7C7
      colorDisabled   =   "frmStockTake1.frx":1D7F1
      colorPressed    =   "frmStockTake1.frx":1D81B
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
      textCaption     =   "frmStockTake1.frx":1D845
      textLT          =   "frmStockTake1.frx":1D8A7
      textCT          =   "frmStockTake1.frx":1D8BF
      textRT          =   "frmStockTake1.frx":1D8D7
      textLM          =   "frmStockTake1.frx":1D8EF
      textRM          =   "frmStockTake1.frx":1D907
      textLB          =   "frmStockTake1.frx":1D91F
      textCB          =   "frmStockTake1.frx":1D937
      textRB          =   "frmStockTake1.frx":1D94F
      colorBack       =   "frmStockTake1.frx":1D967
      colorIntern     =   "frmStockTake1.frx":1D991
      colorMO         =   "frmStockTake1.frx":1D9BB
      colorFocus      =   "frmStockTake1.frx":1D9E5
      colorDisabled   =   "frmStockTake1.frx":1DA0F
      colorPressed    =   "frmStockTake1.frx":1DA39
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
      textCaption     =   "frmStockTake1.frx":1DA63
      textLT          =   "frmStockTake1.frx":1DAC5
      textCT          =   "frmStockTake1.frx":1DADD
      textRT          =   "frmStockTake1.frx":1DAF5
      textLM          =   "frmStockTake1.frx":1DB0D
      textRM          =   "frmStockTake1.frx":1DB25
      textLB          =   "frmStockTake1.frx":1DB3D
      textCB          =   "frmStockTake1.frx":1DB55
      textRB          =   "frmStockTake1.frx":1DB6D
      colorBack       =   "frmStockTake1.frx":1DB85
      colorIntern     =   "frmStockTake1.frx":1DBAF
      colorMO         =   "frmStockTake1.frx":1DBD9
      colorFocus      =   "frmStockTake1.frx":1DC03
      colorDisabled   =   "frmStockTake1.frx":1DC2D
      colorPressed    =   "frmStockTake1.frx":1DC57
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
      textCaption     =   "frmStockTake1.frx":1DC81
      textLT          =   "frmStockTake1.frx":1DCE3
      textCT          =   "frmStockTake1.frx":1DCFB
      textRT          =   "frmStockTake1.frx":1DD13
      textLM          =   "frmStockTake1.frx":1DD2B
      textRM          =   "frmStockTake1.frx":1DD43
      textLB          =   "frmStockTake1.frx":1DD5B
      textCB          =   "frmStockTake1.frx":1DD73
      textRB          =   "frmStockTake1.frx":1DD8B
      colorBack       =   "frmStockTake1.frx":1DDA3
      colorIntern     =   "frmStockTake1.frx":1DDCD
      colorMO         =   "frmStockTake1.frx":1DDF7
      colorFocus      =   "frmStockTake1.frx":1DE21
      colorDisabled   =   "frmStockTake1.frx":1DE4B
      colorPressed    =   "frmStockTake1.frx":1DE75
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
      textCaption     =   "frmStockTake1.frx":1DE9F
      textLT          =   "frmStockTake1.frx":1DF01
      textCT          =   "frmStockTake1.frx":1DF19
      textRT          =   "frmStockTake1.frx":1DF31
      textLM          =   "frmStockTake1.frx":1DF49
      textRM          =   "frmStockTake1.frx":1DF61
      textLB          =   "frmStockTake1.frx":1DF79
      textCB          =   "frmStockTake1.frx":1DF91
      textRB          =   "frmStockTake1.frx":1DFA9
      colorBack       =   "frmStockTake1.frx":1DFC1
      colorIntern     =   "frmStockTake1.frx":1DFEB
      colorMO         =   "frmStockTake1.frx":1E015
      colorFocus      =   "frmStockTake1.frx":1E03F
      colorDisabled   =   "frmStockTake1.frx":1E069
      colorPressed    =   "frmStockTake1.frx":1E093
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
      textCaption     =   "frmStockTake1.frx":1E0BD
      textLT          =   "frmStockTake1.frx":1E11F
      textCT          =   "frmStockTake1.frx":1E137
      textRT          =   "frmStockTake1.frx":1E14F
      textLM          =   "frmStockTake1.frx":1E167
      textRM          =   "frmStockTake1.frx":1E17F
      textLB          =   "frmStockTake1.frx":1E197
      textCB          =   "frmStockTake1.frx":1E1AF
      textRB          =   "frmStockTake1.frx":1E1C7
      colorBack       =   "frmStockTake1.frx":1E1DF
      colorIntern     =   "frmStockTake1.frx":1E209
      colorMO         =   "frmStockTake1.frx":1E233
      colorFocus      =   "frmStockTake1.frx":1E25D
      colorDisabled   =   "frmStockTake1.frx":1E287
      colorPressed    =   "frmStockTake1.frx":1E2B1
      Orientation     =   3
      TextCaptionAlignment=   0
      HollowFrame     =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdTot 
      Height          =   1350
      Index           =   0
      Left            =   390
      TabIndex        =   60
      Top             =   2280
      Width           =   2265
      _Version        =   524298
      _ExtentX        =   3995
      _ExtentY        =   2381
      _StockProps     =   66
      Caption         =   "Tot"
      Enabled         =   0   'False
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
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      textCaption     =   "frmStockTake1.frx":1E2DB
      textLT          =   "frmStockTake1.frx":1E341
      textCT          =   "frmStockTake1.frx":1E359
      textRT          =   "frmStockTake1.frx":1E371
      textLM          =   "frmStockTake1.frx":1E389
      textRM          =   "frmStockTake1.frx":1E3A1
      textLB          =   "frmStockTake1.frx":1E3B9
      textCB          =   "frmStockTake1.frx":1E3D1
      textRB          =   "frmStockTake1.frx":1E3E9
      colorBack       =   "frmStockTake1.frx":1E401
      colorIntern     =   "frmStockTake1.frx":1E42B
      colorMO         =   "frmStockTake1.frx":1E455
      colorFocus      =   "frmStockTake1.frx":1E47F
      colorDisabled   =   "frmStockTake1.frx":1E4A9
      colorPressed    =   "frmStockTake1.frx":1E4D3
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTot 
      Height          =   1350
      Index           =   1
      Left            =   2655
      TabIndex        =   61
      Top             =   2280
      Width           =   2265
      _Version        =   524298
      _ExtentX        =   3995
      _ExtentY        =   2381
      _StockProps     =   66
      Caption         =   "750ml"
      Enabled         =   0   'False
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
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      textCaption     =   "frmStockTake1.frx":1E4FD
      textLT          =   "frmStockTake1.frx":1E567
      textCT          =   "frmStockTake1.frx":1E57F
      textRT          =   "frmStockTake1.frx":1E597
      textLM          =   "frmStockTake1.frx":1E5AF
      textRM          =   "frmStockTake1.frx":1E5C7
      textLB          =   "frmStockTake1.frx":1E5DF
      textCB          =   "frmStockTake1.frx":1E5F7
      textRB          =   "frmStockTake1.frx":1E60F
      colorBack       =   "frmStockTake1.frx":1E627
      colorIntern     =   "frmStockTake1.frx":1E651
      colorMO         =   "frmStockTake1.frx":1E67B
      colorFocus      =   "frmStockTake1.frx":1E6A5
      colorDisabled   =   "frmStockTake1.frx":1E6CF
      colorPressed    =   "frmStockTake1.frx":1E6F9
      Orientation     =   1
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdCorr 
      Height          =   1350
      Left            =   4920
      TabIndex        =   62
      Top             =   2280
      Width           =   2520
      _Version        =   524298
      _ExtentX        =   4445
      _ExtentY        =   2381
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
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      HighlightColor  =   192
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmStockTake1.frx":1E723
      textLT          =   "frmStockTake1.frx":1E799
      textCT          =   "frmStockTake1.frx":1E7B1
      textRT          =   "frmStockTake1.frx":1E7C9
      textLM          =   "frmStockTake1.frx":1E7E1
      textRM          =   "frmStockTake1.frx":1E7F9
      textLB          =   "frmStockTake1.frx":1E811
      textCB          =   "frmStockTake1.frx":1E829
      textRB          =   "frmStockTake1.frx":1E841
      colorBack       =   "frmStockTake1.frx":1E859
      colorIntern     =   "frmStockTake1.frx":1E883
      colorMO         =   "frmStockTake1.frx":1E8AD
      colorFocus      =   "frmStockTake1.frx":1E8D7
      colorDisabled   =   "frmStockTake1.frx":1E901
      colorPressed    =   "frmStockTake1.frx":1E92B
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
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
      textCaption     =   "frmStockTake1.frx":1E955
      textLT          =   "frmStockTake1.frx":1E9B7
      textCT          =   "frmStockTake1.frx":1E9CF
      textRT          =   "frmStockTake1.frx":1E9E7
      textLM          =   "frmStockTake1.frx":1E9FF
      textRM          =   "frmStockTake1.frx":1EA17
      textLB          =   "frmStockTake1.frx":1EA2F
      textCB          =   "frmStockTake1.frx":1EA47
      textRB          =   "frmStockTake1.frx":1EA5F
      colorBack       =   "frmStockTake1.frx":1EA77
      colorIntern     =   "frmStockTake1.frx":1EAA1
      colorMO         =   "frmStockTake1.frx":1EACB
      colorFocus      =   "frmStockTake1.frx":1EAF5
      colorDisabled   =   "frmStockTake1.frx":1EB1F
      colorPressed    =   "frmStockTake1.frx":1EB49
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
      TabIndex        =   64
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
      textCaption     =   "frmStockTake1.frx":1EB73
      textLT          =   "frmStockTake1.frx":1EBD5
      textCT          =   "frmStockTake1.frx":1EBED
      textRT          =   "frmStockTake1.frx":1EC05
      textLM          =   "frmStockTake1.frx":1EC1D
      textRM          =   "frmStockTake1.frx":1EC35
      textLB          =   "frmStockTake1.frx":1EC4D
      textCB          =   "frmStockTake1.frx":1EC65
      textRB          =   "frmStockTake1.frx":1EC7D
      colorBack       =   "frmStockTake1.frx":1EC95
      colorIntern     =   "frmStockTake1.frx":1ECBF
      colorMO         =   "frmStockTake1.frx":1ECE9
      colorFocus      =   "frmStockTake1.frx":1ED13
      colorDisabled   =   "frmStockTake1.frx":1ED3D
      colorPressed    =   "frmStockTake1.frx":1ED67
      Style           =   2
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   2
      Left            =   1710
      TabIndex        =   72
      Top             =   525
      Width           =   1245
      _Version        =   524298
      _ExtentX        =   2196
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Finalize Take"
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
      textCaption     =   "frmStockTake1.frx":1ED91
      textLT          =   "frmStockTake1.frx":1EE0B
      textCT          =   "frmStockTake1.frx":1EE23
      textRT          =   "frmStockTake1.frx":1EE3B
      textLM          =   "frmStockTake1.frx":1EE53
      textRM          =   "frmStockTake1.frx":1EE6B
      textLB          =   "frmStockTake1.frx":1EE83
      textCB          =   "frmStockTake1.frx":1EE9B
      textRB          =   "frmStockTake1.frx":1EEB3
      colorBack       =   "frmStockTake1.frx":1EECB
      colorIntern     =   "frmStockTake1.frx":1EEF5
      colorMO         =   "frmStockTake1.frx":1EF1F
      colorFocus      =   "frmStockTake1.frx":1EF49
      colorDisabled   =   "frmStockTake1.frx":1EF73
      colorPressed    =   "frmStockTake1.frx":1EF9D
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin btButtonEx.ButtonEx cmdInput 
      Height          =   1080
      Index           =   12
      Left            =   10830
      TabIndex        =   84
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
      Caption         =   "<All Locations>"
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
      textCaption     =   "frmStockTake1.frx":1EFC7
      textLT          =   "frmStockTake1.frx":1F045
      textCT          =   "frmStockTake1.frx":1F05D
      textRT          =   "frmStockTake1.frx":1F075
      textLM          =   "frmStockTake1.frx":1F08D
      textRM          =   "frmStockTake1.frx":1F0A5
      textLB          =   "frmStockTake1.frx":1F0BD
      textCB          =   "frmStockTake1.frx":1F0D5
      textRB          =   "frmStockTake1.frx":1F0ED
      colorBack       =   "frmStockTake1.frx":1F105
      colorIntern     =   "frmStockTake1.frx":1F12F
      colorMO         =   "frmStockTake1.frx":1F159
      colorFocus      =   "frmStockTake1.frx":1F183
      colorDisabled   =   "frmStockTake1.frx":1F1AD
      colorPressed    =   "frmStockTake1.frx":1F1D7
      Orientation     =   1
      TextCaptionAlignment=   0
      HollowFrame     =   -1  'True
   End
   Begin VB.Label lblWeight_Empty 
      Height          =   285
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblWeight_Full 
      Height          =   285
      Left            =   3060
      TabIndex        =   82
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label QAnswer 
      Height          =   345
      Left            =   1680
      TabIndex        =   73
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSForms.Label lblDepart 
      Height          =   285
      Left            =   6420
      TabIndex        =   63
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
      Picture         =   "frmStockTake1.frx":1F201
   End
End
Attribute VB_Name = "frmStockTake1"
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
                cmdPlu(i).TextDescrCB.Text = "Not Counted"
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
Private Sub BtnEnh1_Click(index As Integer)
    If picSlip.Visible = True Then Exit Sub
    If Timer2.Enabled = True Then Exit Sub
    Select Case grdLoc.Visible
         Case True
            grdLoc.Visible = False
         Case False
            grdLoc.Rows = 0
            grdLoc.Visible = True
            ActiveReadServer "Select *,isnull((Select count(Location_No) from Stock_Take_Listing where Stock_Take_Listing.Location_No = Locations.Location_No group by Location_No),0)as Products from Locations where Stock_Take = 0 order by Location_No"
            While Not rs.EOF
                grdLoc.Rows = grdLoc.Rows + 1
                grdLoc.TextMatrix(grdLoc.Rows - 1, 0) = rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                Select Case rs.Fields("Products")
                    Case 0: grdLoc.TextMatrix(grdLoc.Rows - 1, 1) = "No Active Count"
                    Case Else: grdLoc.TextMatrix(grdLoc.Rows - 1, 1) = "Active Count"
                End Select
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
Private Sub BtnEnh1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub
Private Sub cmdArrow_Click(index As Integer)
    Select Case index
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
Private Sub cmdArrow_MouseDown(index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    scrolTimer.Tag = index
    scrolTimer.Enabled = True
End Sub

Private Sub cmdArrow1_Click(index As Integer)
    Select Case index
        Case 0
            If grdMerge.Row <> 1 Then
                grdMerge.Row = grdMerge.Row - 1
            End If
        Case 1
            If grdMerge.Row <> grdMerge.Rows - 1 Then
                grdMerge.Row = grdMerge.Row + 1
            End If
    End Select
    grdMerge.ShowCell grdMerge.Row, 0
End Sub

Private Sub cmdArrow1_MouseDown(index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    scrolTimer1.Tag = index
    scrolTimer1.Enabled = True
End Sub
Private Sub cmdArrow1_MouseUp(index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    scrolTimer1.Enabled = False
    scrolTimer1.Interval = 700
End Sub

Private Sub cmdClose1_Click()
    picMerge.Visible = False
End Sub

Private Sub cmdCorr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub
Private Sub cmdDeptStrip_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub

Private Sub cmdErr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub
Private Sub cmdFancy_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
    picHoldFocus.SetFocus
End Sub
Private Sub cmdInput_KeyDown(index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
    Me.SetFocus
End Sub
Private Sub cmdLogOff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub
Private Sub cmdPlu_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
    picHoldFocus.SetFocus
End Sub
Private Sub cmdPrint_Click()
    QAnswer.Caption = ""
    Load frmQuestion
    frmQuestion.Tag = "Take Print"
    frmStockTake.Tag = "Not Now"
    frmQuestion.lblCap = "Please Select a Print Layout for the Variance Report"
    frmQuestion.Show vbModal
    Select Case QAnswer
        Case "A4"
            ActiveUpdateServer "Delete from Stock_Take_Temp"
            DoEvents
            For i = 1 To grdMerge.Rows - 1
                ActiveReadServer "Select Ave_Cost from Products where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "'"
                Ave_Cost = 0
                If rs.RecordCount > 0 Then
                    Ave_Cost = Val(rs.Fields("Ave_Cost") & "")
                End If
                rs.Close
                Variance_Val = grdMerge.ValueMatrix(i, 6) * Ave_Cost
                Stock_Value = grdMerge.ValueMatrix(i, 7) * Ave_Cost
                ActiveUpdateServer "INSERT INTO [Stock_Take_Temp]([Product_code], [Department_No], [Location_No], [Qty_Counted], [Qty_on_Hand], [Ave_Cost], [Description], [Variance_Qty], [Variance_Value],[Stock_Value])" & _
                " Values ('" & grdMerge.TextMatrix(i, 4) & "','" & grdMerge.TextMatrix(i, 5) & "','" & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & "','" & grdMerge.TextMatrix(i, 2) & "','" & grdMerge.TextMatrix(i, 1) & "','" & Ave_Cost & "','" & grdMerge.TextMatrix(i, 0) & "','" & grdMerge.TextMatrix(i, 3) & "','" & Variance_Val & "','" & Stock_Value & "')"
            Next i
            rptVarSheet.Show
            Exit Sub
    End Select
    On Error GoTo trap
    Screen.MousePointer = 11
    frmTillReport.Tag = "Not Now"
    PrintErr = 0
    Slip_Port = ""
    filenum = FreeFile
    Close #filenum
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
    Print #filenum, "STOCK VARIANCE REPORT"
    Print #filenum, lblDate.Caption
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, UCase(cmdUser.Caption)
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    CurrentDepart = ""
    '
    For i = 1 To grdMerge.Rows - 1
        If CurrentDepart <> grdMerge.TextMatrix(i, 5) Then
            CurrentDepart = grdMerge.TextMatrix(i, 5)
            Print #filenum, Chr(27) & Chr(97) & Chr(49);
            Print #filenum, Chr(27) & Chr(45) & Chr(49);
            Print #filenum, Chr(27) & Chr(50);
            Print #filenum, UCase(grdMerge.TextMatrix(i, 5))
            Print #filenum, Chr(27) & Chr(97) & Chr(48);
            Print #filenum, Chr(27) & Chr(45) & Chr(48);
        End If
        Print #filenum, Chr(27) & Chr(97) & Chr(48);
        Print #filenum, grdMerge.TextMatrix(i, 0)
        Print #filenum, Chr(27) & Chr(97) & Chr(50);
        Print #filenum, Chr(27) & Chr(51) & Chr(18);
        Print #filenum, Chr(218) & String(11, Chr(196)) & Chr(191) & Chr(218) & String(11, Chr(196)) & Chr(191) & Chr(218) & String(11, Chr(196)) & Chr(191)
        SOH = ""
        tSOH = String(11, " ")
        Counted = ""
        tCounted = String(11, " ")
        Variance = ""
        tVariance = String(11, " ")
        SOH = Replace(grdMerge.TextMatrix(i, 1), "(Tots)", "") & String(11 - Len(Replace(grdMerge.TextMatrix(i, 1), "(Tots)", "")), Chr(32))
        If InStr(grdMerge.TextMatrix(i, 1), "(Tots)") <> 0 Then
            tSOH = "(Tots)     "
        End If
        Counted = Replace(grdMerge.TextMatrix(i, 2), "(Tots)", "") & String(11 - Len(Replace(grdMerge.TextMatrix(i, 2), "(Tots)", "")), Chr(32))
        If InStr(grdMerge.TextMatrix(i, 2), "(Tots)") <> 0 Then
            tCounted = "(Tots)     "
        End If
        Variance = Replace(grdMerge.TextMatrix(i, 3), "(Tots)", "") & String(11 - Len(Replace(grdMerge.TextMatrix(i, 3), "(Tots)", "")), Chr(32))
        If InStr(grdMerge.TextMatrix(i, 3), "(Tots)") <> 0 Then
            tVariance = "(Tots)     "
        End If
        Print #filenum, Chr(179) & SOH & Chr(179) & Chr(179) & Counted & Chr(179) & Chr(179) & Variance & Chr(179)
        Print #filenum, Chr(179) & tSOH & Chr(179) & Chr(179) & tCounted & Chr(179) & Chr(179) & tVariance & Chr(179)
        Print #filenum, Chr(192) & String(11, Chr(196)) & Chr(217) & Chr(192) & String(11, Chr(196)) & Chr(217) & Chr(192) & String(11, Chr(196)) & Chr(217)
    Next i
    Print #filenum, Chr(27) & Chr(50);
    Print #filenum, String(40, "-")
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(51) & Chr(18);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    TotalCounted = 0
    TotalToCount = 0
    ActiveReadServer "Select * from Departments_Stock where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & " and Department_No = '" & Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) & "'"
    If rs.RecordCount > 0 Then
        TotalCounted = rs.Fields("Product_Count") & String(12 - Len(rs.Fields("Product_Count")), Chr(32))
        TotalToCount = rs.Fields("Product_List") & String(12 - Len(rs.Fields("Product_List")), Chr(32))
        Difference = rs.Fields("Product_List") - rs.Fields("Product_Count") & String(12 - Len(rs.Fields("Product_List") - rs.Fields("Product_Count")), Chr(32))
    End If
    rs.Close
    
    Print #filenum, Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, "TOTAL PRODUCTS:  " & Chr(179) & TotalToCount & Chr(179)
    Print #filenum, Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, "TOTAL COUNTED:  " & Chr(179) & TotalCounted & Chr(179)
    Print #filenum, Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(12, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(12, Chr(32)) & Chr(179)
    Print #filenum, "DIFFERENCE:  " & Chr(179) & Difference & Chr(179)
    Print #filenum, Chr(192) & String(12, Chr(196)) & Chr(217)
    
    Print #filenum, String(40, "-")
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
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(100) & Chr(7);
    Print #filenum, Chr(29) & "V" & Chr(49);
    Print #filenum, Chr(27) & Chr(64);
    Print #filenum, Chr(27) & Chr(69) & Chr(1);
    Close #1
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim X As Printer
        For Each X In Printers
            If UCase(X.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = X.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmStockTake.Tag = "Not Now"
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = Err.Description
    DoEvents
    frmError.Show vbModal
    Screen.MousePointer = 0
    On Error GoTo 0
End Sub

Private Sub cmdRemove_Click()
    If grdRem.Rows = 1 Then Exit Sub
    ActiveUpdateServer "Delete from Stock_Take_Listing where Line_No = " & grdRem.TextMatrix(grdRem.Row, 3)
    grdRem.RemoveItem (grdRem.Row)
    grdMain.RemoveItem (grdMain.Row)
    ActiveReadServer "Select * from Departments_Stock where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & " and Department_No = '" & Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) & "'"
    If rs.RecordCount > 0 Then
        cmdDeptStrip(Val(lblDepart.Tag)).TextDescrCB.Text = rs.Fields("Product_Count") & " Counted of " & rs.Fields("Product_List")
        For i = 0 To grdDept.Rows - 1
            If grdDept.TextMatrix(i, 1) = Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) Then
                grdDept.TextMatrix(i, 3) = rs.Fields("Product_Count") & " Counted of " & rs.Fields("Product_List")
            End If
        Next i
    End If
    rs.Close
End Sub

Private Sub cmdTot_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub

Private Sub cmdView_Click()
    frmStockTake.Tag = "Not Now"
    frmStockCount.Show vbModal
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
            Take_No = 0
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
            lblDetail.Caption = "Counted: " & Qty_Counted & " x " & rs.Fields("Short_Description")
            
            ActiveUpdateServer "INSERT INTO Stock_Take_Listing (Date_Time,  Workstation_No, User_No, Location_No,Product_Code,Department_No,Qty_Counted,Unit_of_Measure) values " & _
            "(Getdate()," & Workstation_No & "," & UserRecord.User_Number & "," & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & ",'" & Product_Code & "','" & rs.Fields("Department_No") & "'," & Qty_Counted & ",'" & Unit_of_Measure & "')"
            
            ActiveReadServer1 "Select max(Line_No) as Line_No from Stock_Take_Listing"
            grdMain.TextMatrix(grdMain.Row, 3) = rs1.Fields("Line_No")
            rs1.Close
            ActiveReadServer1 "Select * from Departments_Stock where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & " and Department_No = '" & rs.Fields("Department_No") & "'"
            If rs1.RecordCount > 0 Then
                cmdDeptStrip(Val(lblDepart.Tag)).TextDescrCB.Text = rs1.Fields("Product_Count") & " Counted of " & rs1.Fields("Product_List")
                For i = 0 To grdDept.Rows - 1
                    If grdDept.TextMatrix(i, 1) = rs.Fields("Department_No") Then
                        grdDept.TextMatrix(i, 3) = rs1.Fields("Product_Count") & " Counted of " & rs1.Fields("Product_List")
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
    If cmdUser.Caption = "<All Locations>" Then
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
                            If Left(lblDetail.Caption, 7) = "Counted" Then
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
Private Sub cmdCorr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblDepart.Tag = "" Then
        cmdErr.Caption = "Select a Department First"
        Timer2.Enabled = True
        cmdErr.Visible = True
        Exit Sub
    End If
    If grdMain.Rows = 1 Then
        cmdErr.Caption = "Invalid Key Pressed"
        Timer2.Enabled = True
        cmdErr.Visible = True
        Exit Sub
    Else
        picSlip.Visible = True
    End If
    lblHead.Caption = lblDepart.Caption
    grdRem.Rows = grdMain.Rows
    For i = 1 To grdMain.Rows - 1
        grdRem.TextMatrix(i, 0) = grdMain.TextMatrix(i, 0)
        grdRem.TextMatrix(i, 1) = grdMain.TextMatrix(i, 1)
        grdRem.TextMatrix(i, 2) = grdMain.TextMatrix(i, 2)
        grdRem.TextMatrix(i, 3) = grdMain.TextMatrix(i, 3)
    Next i
    grdRem.Row = grdMain.Row
    grdRem.ShowCell grdRem.Row, 0
End Sub
Private Sub cmdDeptStrip_Click(index As Integer)
    Me.SetFocus
    DoEvents
    If picSlip.Visible = True Then
        cmdDeptStrip(index).Value = 0
        Exit Sub
    End If
    cmdTot(0).Enabled = False
    cmdTot(1).Enabled = False
    If cmdDeptStrip(index).Picture = App.Path & "\icons\downArr.bmp" Then
        cmdDeptStrip(index).Value = 0
        cmdDeptStrip(index).TextDescrCB.ColorNormal = vbYellow
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
    If cmdDeptStrip(index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdDeptStrip(index).Value = 0
        cmdDeptStrip(index).TextDescrCB.ColorNormal = vbYellow
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

Private Sub cmdDeptStrip_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 15
        If index <> i Then
            cmdDeptStrip(i).TextDescrCB.ColorNormal = vbYellow
        Else
            cmdDeptStrip(i).TextDescrCB.ColorNormal = &HC0&
        End If
    Next i
    Load_List cmdDeptStrip(index).Tag, Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
    If picMerge.Visible = True Then
        Merge_Take
    End If
    If cmdDeptStrip(index).Picture <> App.Path & "\icons\downArr.bmp" And cmdDeptStrip(index).Picture <> App.Path & "\icons\upArr.bmp" Then
        lblDepart.Caption = "Counting: " & cmdDeptStrip(index).Tag & " > " & Replace(cmdDeptStrip(index).Caption, "&&", "&")
        lblDepart.Tag = index
        LoadPlu cmdDeptStrip(index).Tag
    End If
    DoEvents
End Sub

Private Sub cmdErr_Click()
    cmdErr.Caption = ""
    Timer2.Enabled = False
    cmdErr.Visible = False
End Sub
Private Sub Merge_Take()
    If grdMain.Rows = 1 Then
        grdMerge.Rows = 1
        Exit Sub
    End If
    Screen.MousePointer = 11
    grdMerge.Rows = 1
    Select Case lblDepart.Tag
        Case ""
        Case Else
    End Select
    For i = 1 To grdMain.Rows - 1
        grdMerge.Rows = grdMerge.Rows + 1
        grdMerge.TextMatrix(i, 0) = grdMain.TextMatrix(i, 0)
        grdMerge.TextMatrix(i, 2) = grdMain.TextMatrix(i, 1)
        grdMerge.TextMatrix(i, 1) = 0
        unit_Size = ""
        Unit_of_Measure = ""
        For b = 1 To Len(grdMerge.TextMatrix(i, 2))
            If Asc(Mid(grdMerge.TextMatrix(i, 2), b, 1)) < 46 Or Asc(Mid(grdMerge.TextMatrix(i, 2), b, 1)) > 57 Then
                Unit_of_Measure = Trim(Mid(grdMerge.TextMatrix(i, 2), b))
                Exit For
            Else
                unit_Size = unit_Size & Mid(grdMerge.TextMatrix(i, 2), b, 1)
            End If
        Next b
        Select Case Unit_of_Measure
            Case "(Tots)"
               Bottle_Size = ""
               Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                For b = 1 To Len(Bottle)
                    If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                        Exit For
                    Else
                        Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                    End If
                Next b
                If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                If Bottle_Size = "1000" Then
                    unit_Size = Round((unit_Size * 40) / Val(Bottle_Size), 4)
                Else
                    unit_Size = Round((unit_Size * 25) / Val(Bottle_Size), 4)
                End If
                grdMerge.TextMatrix(i, 2) = unit_Size
            Case "ml"
                Bottle_Size = ""
                Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                For b = 1 To Len(Bottle)
                    If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                        Exit For
                    Else
                        Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                    End If
                Next b
                If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
                grdMerge.TextMatrix(i, 2) = unit_Size
            Case "kg"
                Bottle_Size = ""
                Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                For b = 1 To Len(Bottle)
                    If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                        Exit For
                    Else
                        Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                    End If
                Next b
                If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
                grdMerge.TextMatrix(i, 2) = unit_Size
            Case "g"
                Bottle_Size = ""
                Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                For b = 1 To Len(Bottle)
                    If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                        Exit For
                    Else
                        Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                    End If
                Next b
                If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
                grdMerge.TextMatrix(i, 2) = unit_Size
            Case "lt"
                Bottle_Size = ""
                Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                For b = 1 To Len(Bottle)
                    If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                        Exit For
                    Else
                        Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                    End If
                Next b
                If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
                grdMerge.TextMatrix(i, 2) = unit_Size
        End Select
        grdMerge.TextMatrix(i, 4) = grdMain.TextMatrix(i, 2)
        
    Next i
    grdMerge.Select 1, 4, grdMerge.Rows - 1, 4
    grdMerge.Sort = flexSortNumericAscending
Restart:
    For i = 1 To grdMerge.Rows - 1
        If i > 1 Then
            If grdMerge.TextMatrix(i, 4) = grdMerge.TextMatrix(i - 1, 4) Then
                grdMerge.TextMatrix(i - 1, 2) = grdMerge.ValueMatrix(i - 1, 2) + grdMerge.ValueMatrix(i, 2)
                grdMerge.TextMatrix(i - 1, 1) = grdMerge.ValueMatrix(i - 1, 1) + grdMerge.ValueMatrix(i, 1)
                grdMerge.RemoveItem i
                GoTo Restart
            End If
        End If
    Next i
    For i = 1 To grdMerge.Rows - 1
        ActiveReadServer "Select Stock_on_Hand from Quantities where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "' and Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
        If rs.RecordCount > 0 Then
            Stock_on_Hand = rs.Fields("Stock_on_Hand")
        Else
            Stock_on_Hand = 0
        End If
        rs.Close
        ActiveReadServer "Select Department_No from Products where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "'"
        If rs.RecordCount > 0 Then
            grdMerge.TextMatrix(i, 5) = rs.Fields("Department_No")
        End If
        rs.Close
        
        grdMerge.TextMatrix(i, 1) = Stock_on_Hand
        grdMerge.TextMatrix(i, 3) = Round(grdMerge.ValueMatrix(i, 2) - grdMerge.ValueMatrix(i, 1), 4)
        grdMerge.TextMatrix(i, 6) = grdMerge.TextMatrix(i, 3)
        grdMerge.TextMatrix(i, 7) = grdMerge.TextMatrix(i, 2)
        Bottle_Size = ""
        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
        For b = 1 To Len(Bottle)
             If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                 Exit For
             Else
                 Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
             End If
        Next b
        Select Case Bottle
            Case "500ml"
                If Round((500 * (grdMerge.ValueMatrix(i, 1) - Round(grdMerge.ValueMatrix(i, 1), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 1) = Round(grdMerge.ValueMatrix(i, 1), 0)
               Else
                   If Round(grdMerge.ValueMatrix(i, 1), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 1) = Int(grdMerge.ValueMatrix(i, 1)) & " & " & Round(((500 * (grdMerge.ValueMatrix(i, 1) - Int(grdMerge.ValueMatrix(i, 1)))) / 25), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 1) = Round(((500 * (grdMerge.ValueMatrix(i, 1) - Int(grdMerge.ValueMatrix(i, 1)))) / 25), 0) & " (Tots)"
                   End If
               End If
                
                If Round((500 * (grdMerge.ValueMatrix(i, 2) - Round(grdMerge.ValueMatrix(i, 2), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 2) = Round(grdMerge.ValueMatrix(i, 2), 0)
                Else
                   If Round(grdMerge.ValueMatrix(i, 2), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 2) = Int(grdMerge.ValueMatrix(i, 2)) & " & " & Round(((500 * (grdMerge.ValueMatrix(i, 2) - Int(grdMerge.ValueMatrix(i, 2)))) / 25), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 2) = Round(((500 * (grdMerge.ValueMatrix(i, 2) - Int(grdMerge.ValueMatrix(i, 2)))) / 25), 0) & " (Tots)"
                   End If
               End If
               
               If Round((500 * (grdMerge.ValueMatrix(i, 3) - Round(grdMerge.ValueMatrix(i, 3), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 3) = Round(grdMerge.ValueMatrix(i, 3), 0)
               Else
                   If Round(grdMerge.ValueMatrix(i, 3), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 3) = Int(grdMerge.ValueMatrix(i, 3)) & " & " & Round(((500 * (grdMerge.ValueMatrix(i, 3) - Round(grdMerge.ValueMatrix(i, 3), 0))) / 25), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 3) = Round(((500 * (grdMerge.ValueMatrix(i, 3) - Round(grdMerge.ValueMatrix(i, 3), 0))) / 25), 0) & " (Tots)"
                   End If
               End If
           Case "750ml"
                If Round((750 * (grdMerge.ValueMatrix(i, 1) - Round(grdMerge.ValueMatrix(i, 1), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 1) = Round(grdMerge.ValueMatrix(i, 1), 0)
                Else
                   If Round(grdMerge.ValueMatrix(i, 1), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 1) = Int(grdMerge.ValueMatrix(i, 1)) & " & " & Round(((750 * (grdMerge.ValueMatrix(i, 1) - Int(grdMerge.ValueMatrix(i, 1)))) / 25), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 1) = Round((750 / 25) * grdMerge.ValueMatrix(i, 1), 0) & " (Tots)"
                   End If
                End If
                
                If Round((750 * (grdMerge.ValueMatrix(i, 2) - Round(grdMerge.ValueMatrix(i, 2), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 2) = Round(grdMerge.ValueMatrix(i, 2), 0)
                Else
                   If Round(grdMerge.ValueMatrix(i, 2), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 2) = Int(grdMerge.ValueMatrix(i, 2)) & " & " & Round(((750 * (grdMerge.ValueMatrix(i, 2) - Int(grdMerge.ValueMatrix(i, 2)))) / 25), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 2) = Round(((750 * (grdMerge.ValueMatrix(i, 2) - Int(grdMerge.ValueMatrix(i, 2)))) / 25), 0) & " (Tots)"
                   End If
               End If
               
               If Round((750 * (grdMerge.ValueMatrix(i, 3) - Round(grdMerge.ValueMatrix(i, 3), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 3) = Round(grdMerge.ValueMatrix(i, 3), 0)
               Else
                   If Round(grdMerge.ValueMatrix(i, 3), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 3) = Int(Mid(grdMerge.TextMatrix(i, 3), 1, InStr(grdMerge.TextMatrix(i, 3), ".") - 1)) & " & " & Round((750 / 25) * grdMerge.ValueMatrix(i, 3), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 3) = Round((750 / 25) * grdMerge.ValueMatrix(i, 3), 0) & " (Tots)"
                   End If
               End If
            Case "1000ml"
                If Round((1000 * (grdMerge.ValueMatrix(i, 1) - Round(grdMerge.ValueMatrix(i, 1), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 1) = Round(grdMerge.ValueMatrix(i, 1), 0)
               Else
                   If Round(grdMerge.ValueMatrix(i, 1), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 1) = Int(grdMerge.ValueMatrix(i, 1)) & " & " & Round(((1000 * (grdMerge.ValueMatrix(i, 1) - Int(grdMerge.ValueMatrix(i, 1)))) / 25), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 1) = Round(((1000 * (grdMerge.ValueMatrix(i, 1) - Int(grdMerge.ValueMatrix(i, 1)))) / 25), 0) & " (Tots)"
                   End If
               End If
            
               If Round((1000 * (grdMerge.ValueMatrix(i, 2) - Round(grdMerge.ValueMatrix(i, 2), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 2) = Round(grdMerge.ValueMatrix(i, 2), 0)
               Else
                   If Round(grdMerge.ValueMatrix(i, 2), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 2) = Int(grdMerge.ValueMatrix(i, 2)) & " & " & Round(((1000 * (grdMerge.ValueMatrix(i, 2) - Int(grdMerge.ValueMatrix(i, 2)))) / 25), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 2) = Round(((1000 * (grdMerge.ValueMatrix(i, 2) - Int(grdMerge.ValueMatrix(i, 2)))) / 25), 0) & " (Tots)"
                   End If
               End If
               
               If Round((1000 * (grdMerge.ValueMatrix(i, 3) - Round(grdMerge.ValueMatrix(i, 3), 0))) / 25, 0) = 0 Then
                   grdMerge.TextMatrix(i, 3) = Round(grdMerge.ValueMatrix(i, 3), 0)
               Else
                   If Round(grdMerge.ValueMatrix(i, 3), 0) <> 0 Then
                       grdMerge.TextMatrix(i, 3) = Int(grdMerge.ValueMatrix(i, 3)) & " & " & Round(((1000 * (grdMerge.ValueMatrix(i, 3) - Round(grdMerge.ValueMatrix(i, 3), 0))) / 25), 0) & " (Tots)"
                   Else
                       grdMerge.TextMatrix(i, 3) = Round(((1000 * (grdMerge.ValueMatrix(i, 3) - Int(grdMerge.ValueMatrix(i, 3)))) / 25), 0) & " (Tots)"
                   End If
               End If
        End Select
    Next i
    grdMerge.Select 1, 0, grdMerge.Rows - 1, 0
    grdMerge.Sort = flexSortStringAscending
    
    grdMerge.Row = grdMerge.Rows - 1
    grdMerge.ShowCell grdMerge.Row, 0
    
    Screen.MousePointer = 0
End Sub
Private Sub Finalize_Take()
    If grdMain.Rows = 1 Then Exit Sub
    QAnswer.Caption = ""
    Load frmQuestion
    frmQuestion.Tag = "Take"
    frmStockTake.Tag = "Not Now"
    frmQuestion.lblCap = "Are you sure you want Finalize this Stock Take?"
    frmQuestion.Show vbModal
    Select Case QAnswer.Caption
        Case "Yes"
            Screen.MousePointer = 11
            For i = 0 To cmdDeptStrip.Count - 1
                cmdDeptStrip(i).Value = 0
                cmdDeptStrip(i).TextDescrCB.ColorNormal = vbYellow
            Next i
            Load_List 0, Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
            Merge_Take
            ActiveReadServer1 "Select isnull(max(Take_No),0)+1 as Take_No from Stock_Take_Journal"
            Take_No = rs1.Fields("Take_No")
            rs1.Close
            For i = 1 To grdMerge.Rows - 1
                unit_Size = ""
                Unit_of_Measure = ""
                If InStr(grdMerge.TextMatrix(i, 2), "&") <> 0 Then
                    Full_Units = Val(Mid(grdMerge.TextMatrix(i, 2), 1, InStr(grdMerge.TextMatrix(i, 2), "&") - 1))
                    grdMerge.TextMatrix(i, 2) = Trim(Mid(grdMerge.TextMatrix(i, 2), InStr(grdMerge.TextMatrix(i, 2), "&") + 1))
                Else
                    Full_Units = 0
                End If
                For b = 1 To Len(grdMerge.TextMatrix(i, 2))
                    If Asc(Mid(grdMerge.TextMatrix(i, 2), b, 1)) < 46 Or Asc(Mid(grdMerge.TextMatrix(i, 2), b, 1)) > 57 Then
                        Unit_of_Measure = Trim(Mid(grdMerge.TextMatrix(i, 2), b))
                        Exit For
                    Else
                        unit_Size = unit_Size & Mid(grdMerge.TextMatrix(i, 2), b, 1)
                    End If
                Next b
                Select Case Unit_of_Measure
                    Case "(Tots)"
                       Bottle_Size = ""
                       Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                        For b = 1 To Len(Bottle)
                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                                Exit For
                            Else
                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                            End If
                        Next b
                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                        If Bottle_Size = "1000" Then
                            unit_Size = Round((unit_Size * 40) / Val(Bottle_Size), 4)
                        Else
                            unit_Size = Round((unit_Size * 25) / Val(Bottle_Size), 4)
                        End If
                        grdMerge.TextMatrix(i, 2) = unit_Size
                    Case "ml"
                        Bottle_Size = ""
                        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                        For b = 1 To Len(Bottle)
                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                                Exit For
                            Else
                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                            End If
                        Next b
                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                        unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
                        grdMerge.TextMatrix(i, 2) = unit_Size
                    Case "kg"
                        Bottle_Size = ""
                        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                        For b = 1 To Len(Bottle)
                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                                Exit For
                            Else
                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                            End If
                        Next b
                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                        unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
                        grdMerge.TextMatrix(i, 2) = unit_Size
                    Case "g"
                        Bottle_Size = ""
                        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                        For b = 1 To Len(Bottle)
                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                                Exit For
                            Else
                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                            End If
                        Next b
                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                        unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
                        grdMerge.TextMatrix(i, 2) = unit_Size
                    Case "lt"
                        Bottle_Size = ""
                        Bottle = Mid(grdMerge.TextMatrix(i, 0), InStrRev(grdMerge.TextMatrix(i, 0), " ") + 1)
                        For b = 1 To Len(Bottle)
                            If Asc(Mid(Bottle, b, 1)) < 46 Or Asc(Mid(Bottle, b, 1)) > 57 Then
                                Exit For
                            Else
                                Bottle_Size = Bottle_Size & Mid(Bottle, b, 1)
                            End If
                        Next b
                        If Val(Bottle_Size) = 0 Then Bottle_Size = 1
                        unit_Size = Round(unit_Size / Val(Bottle_Size), 4)
                        grdMerge.TextMatrix(i, 2) = unit_Size
                End Select
                Qty_Count = grdMerge.TextMatrix(i, 2) + Full_Units
                ActiveReadServer "Select Stock_on_Hand from Quantities where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "' and Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
                If rs.RecordCount > 0 Then
                    Stock_on_Hand = rs.Fields("Stock_on_Hand")
                    ActiveUpdateServer "Update Quantities Set Stock_on_Hand = " & Qty_Count & " where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "' and Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
                Else
                    Stock_on_Hand = 0
                    ActiveUpdateServer "INSERT INTO Quantities (Product_Code,Location_No,Stock_on_Hand) values ('" & grdMerge.TextMatrix(i, 4) & "'," & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & "," & Qty_Count & ")"
                End If
                rs.Close
                DoEvents
                ActiveReadServer "Select Ave_Cost from Products where Product_Code = '" & grdMerge.TextMatrix(i, 4) & "'"
                If rs.RecordCount > 0 Then
                    Ave_Cost = rs.Fields("Ave_Cost")
                Else
                    Ave_Cost = 0
                End If
                DoEvents
                ActiveUpdateServer "INSERT INTO Stock_Take_Journal(Take_No, Date_Time, Workstation_No, User_No, Location_No, Product_Code, Qty_on_Hand, Qty_Counted,Ave_Cost)" & _
                " VALUES(" & Take_No & ",Getdate()," & Workstation_No & "," & UserRecord.User_Number & "," & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & ",'" & grdMerge.TextMatrix(i, 4) & "'," & Stock_on_Hand & ",'" & Qty_Count & "'," & Ave_Cost & ")"
                DoEvents
            Next i
            ActiveUpdateServer "Delete from Stock_Take_Listing where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
            DoEvents
            frmStockTake.Tag = ""
            Screen.MousePointer = 11
            Form_Activate
            Screen.MousePointer = 0
            MsgBox "Your Stock Take was Finalized as Take_No = " & Take_No, vbInformation, "HeroPOS"
    End Select
    Screen.MousePointer = 0
    
End Sub
Private Sub cmdFancy_Click(index As Integer)
    If picSlip.Visible = True Then Exit Sub
    Select Case cmdFancy(index).Caption
        Case "Finalize Take"
            If cmdUser.Caption = "<All Locations>" Then
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
        Case "Merge Take"
            Select Case UserRecord.uType
                Case 2, 3, 4, 8
                    cmdErr.Caption = "Access Denied"
                    Timer2.Enabled = True
                    cmdErr.Visible = True
                    Exit Sub
            End Select
            If cmdUser.Caption = "<All Locations>" Then
                cmdErr.Caption = "Select a Location First"
                Timer2.Enabled = True
                cmdErr.Visible = True
                Exit Sub
            End If
            If picMerge.Visible = True Then
                cmdErr.Caption = "Invalid key Pressed"
                Timer2.Enabled = True
                cmdErr.Visible = True
                Exit Sub
            End If
            picMerge.Visible = True
            For i = 0 To cmdDeptStrip.Count - 1
                cmdDeptStrip(i).Value = 0
                cmdDeptStrip(i).TextDescrCB.ColorNormal = vbYellow
            Next i
            Load_List 0, Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
            Merge_Take
        Case "Print Count Sheet"
            If picMerge.Visible = True Then
                cmdErr.Caption = "Invalid key Pressed"
                Timer2.Enabled = True
                cmdErr.Visible = True
                Exit Sub
            End If
            If cmdUser.Caption = "<All Locations>" Then
                cmdErr.Caption = "Select a Location First"
                Timer2.Enabled = True
                cmdErr.Visible = True
                Exit Sub
            End If
            QAnswer.Caption = ""
            Load frmQuestion
            frmQuestion.Tag = "Take Print"
            frmStockTake.Tag = "Not Now"
            frmQuestion.lblCap = "Please Select a Print Layout for the Count Sheet"
            frmQuestion.Show vbModal
            Select Case QAnswer
                Case "A4"
                    If PrintBarStock = 0 Then
                        rptCountSheet.Show
                    Else
                        rptCountSheetBar.Show
                    End If
                Case "Slip"
                    Select Case lblDepart.Tag
                        Case ""
                            Print_Count_Sheet Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1), ""
                        Case Else
                            Print_Count_Sheet Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1), cmdDeptStrip(Val(lblDepart.Tag)).Tag
                    End Select
                Case Else
                    Exit Sub
            End Select
    End Select
    picHoldFocus.SetFocus
End Sub
Private Sub Print_Count_Sheet(Location_No, Department_No)
    On Error GoTo trap
    Screen.MousePointer = 11
    frmTillReport.Tag = "Not Now"
    PrintErr = 0
    Slip_Port = ""
    filenum = FreeFile
    Close #filenum
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
    Print #filenum, "COUNT SHEET"
    Print #filenum, lblDate.Caption
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, UCase(cmdUser.Caption)
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    CurrentDepart = ""
    Select Case Department_No
        Case ""
            ActiveReadServer "Select Department_No,(Select Dept_Name from Departments where Departments.Department_No = Products.Department_No) as Dept_Name,Short_Description, Unit_Size, Unit_of_Measure from Products where Stock_Item = 1 and Department_No in" & _
            " (Select Department_No from Departments_Stock where Location_No = " & Location_No & ") order by Department_No, Description"
        Case Else
            ActiveReadServer "Select Department_No,(Select Dept_Name from Departments where Departments.Department_No = Products.Department_No ) as Dept_Name, Short_Description, Unit_Size, Unit_of_Measure from Products where Stock_Item = 1 and Department_No in" & _
            " (Select Department_No from Departments_Stock where Location_No = " & Location_No & " and Department_No ='" & Department_No & "') order by Department_No, Description"
    End Select
    While Not rs.EOF
        If CurrentDepart <> rs.Fields("Dept_Name") Then
            CurrentDepart = rs.Fields("Dept_Name")
            Print #filenum, Chr(27) & Chr(97) & Chr(49);
            Print #filenum, Chr(27) & Chr(45) & Chr(49);
            Print #filenum, Chr(27) & Chr(50);
            Print #filenum, UCase(rs.Fields("Dept_Name"))
            Print #filenum, Chr(27) & Chr(97) & Chr(48);
            Print #filenum, Chr(27) & Chr(45) & Chr(48);
        End If
        If rs.Fields("Unit_Size") = 0 Then
            unit_Size = rs.Fields("Unit_of_Measure")
        Else
            unit_Size = rs.Fields("Unit_Size") & rs.Fields("Unit_of_Measure")
        End If
        Print #filenum, Chr(27) & Chr(97) & Chr(48);
        Print #filenum, rs.Fields("Short_Description") & " " & unit_Size
        Print #filenum, Chr(27) & Chr(97) & Chr(50);
        Print #filenum, Chr(27) & Chr(51) & Chr(18);
        Print #filenum, Chr(218) & String(13, Chr(196)) & Chr(191) & "  " & Chr(218) & String(13, Chr(196)) & Chr(191)
        Print #filenum, Chr(179) & String(13, Chr(32)) & Chr(179) & "  " & Chr(179) & String(13, Chr(32)) & Chr(179)
        Print #filenum, Chr(179) & String(13, Chr(32)) & Chr(179) & "  " & Chr(179) & String(13, Chr(32)) & Chr(179)
        Print #filenum, Chr(192) & String(13, Chr(196)) & Chr(217) & "  " & Chr(192) & String(13, Chr(196)) & Chr(217)
        rs.MoveNext
    Wend
    rs.Close
    
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
    Print #filenum, Chr(29) & "V" & Chr(49);
    Print #filenum, Chr(27) & Chr(64);
    Print #filenum, Chr(27) & Chr(69) & Chr(1);
    Close #1
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim X As Printer
        For Each X In Printers
            If UCase(X.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = X.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmStockTake.Tag = "Not Now"
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = Err.Description
    DoEvents
    frmError.Show vbModal
    On Error GoTo 0
End Sub
Private Sub cmdInput_MouseDown(index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If picSlip.Visible = True Then Exit Sub
    If picMerge.Visible = True Then Exit Sub
    If cmdUser.Caption = "<All Locations>" Then
        cmdErr.Caption = "Select a Location First"
        Timer2.Enabled = True
        cmdErr.Visible = True
        Exit Sub
    End If
    Select Case cmdInput(index).Caption
        Case "X"
            If InStr(lblDetail.Caption, Chr(215)) <> 0 Then Exit Sub
            lblDetail.Caption = lblDetail.Caption & " " & Chr(215) & " "
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
            lblDetail.Font.Size = 26
            If lblDetail.Caption <> "" Then
                If Asc(Mid(lblDetail.Caption, 1, 1)) < 48 Or Asc(Mid(lblDetail.Caption, 1, 1)) > 57 Then
                    If lblDetail.Caption = "." Then
                       lblDetail.Caption = lblDetail.Caption & cmdInput(index).Caption
                    Else
                        If InStr(lblDetail.Caption, ".") = 0 Then
                            lblDetail.Caption = cmdInput(index).Caption
                        Else
                            If Left(lblDetail.Caption, 7) = "Counted" Then
                                lblDetail.Caption = cmdInput(index).Caption
                            Else
                                lblDetail.Caption = lblDetail.Caption & cmdInput(index).Caption
                            End If
                        End If
                    End If
                Else
                    lblDetail.Caption = lblDetail.Caption & cmdInput(index).Caption
                End If
            Else
                lblDetail.Caption = lblDetail.Caption & cmdInput(index).Caption
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
                lblDetail.Caption = cmdInput(index).Caption
                Exit Sub
            End If
            If Asc(Mid(lblDetail.Caption, 1, 1)) < 48 Or Asc(Mid(lblDetail.Caption, 1, 1)) > 57 Then
                lblDetail.Caption = cmdInput(index).Caption
            Else
                lblDetail.Caption = lblDetail.Caption & cmdInput(index).Caption
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
Private Sub cmdKey_Click(index As Integer)
    grdLoc.Visible = False
    Me.Hide
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
    If cmdUser.Caption = "<All Locations>" Then
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
            ActiveUpdateServer "Delete from Stock_Take_Listing where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
            QAnswer.Caption = ""
            lblDetail.Caption = "Select a Location to Count"
            For i = 0 To cmdDeptStrip.Count - 1
                cmdDeptStrip(i).Visible = False
            Next i
            cmdUser.Caption = "<All Locations>"
            grdMain.Rows = 1
            For i = 0 To cmdPlu.Count - 1
                cmdPlu(i).Visible = False
            Next i
            cmdTot(0).Enabled = False
            cmdTot(1).Enabled = False
            picMerge.Visible = False
        Case "No"
            DoEvents
    End Select
End Sub

Private Sub cmdPlu_Click(index As Integer)
    If picSlip.Visible = True Then
        cmdPlu(index).Value = 0
        Exit Sub
    End If
    DoEvents
    If cmdPlu(index).Picture = App.Path & "\icons\downArr.bmp" Then
        grdPlu.Row = grdPlu.Row + 1
        For i = 0 To 17
            If grdPlu.TextMatrix(grdPlu.Row, 1) = "ßArrow" Then
                If i = 0 Then
                    cmdPlu(i).Caption = ""
                    cmdPlu(i).ToolTipText = ""
                    cmdPlu(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                Else
                    cmdPlu(i).Caption = ""
                    cmdPlu(i).ToolTipText = ""
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
                    cmdPlu(i).TextDescrCB.Text = "Not Counted"
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
        cmdPlu(index).Value = 0
    End If
    If cmdPlu(index).Picture = App.Path & "\icons\upArr.bmp" Then
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
                    cmdPlu(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdPlu(i).Visible = False Then cmdPlu(i).Visible = True
                Else
                    cmdPlu(i).Caption = ""
                     cmdPlu(i).ToolTipText = ""
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
                    cmdPlu(i).TextDescrCB.Text = "Not Counted"
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
        cmdPlu(index).Value = 0
    End If
End Sub

Private Sub cmdPlu_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdPlu(index).Picture <> App.Path & "\icons\downArr.bmp" And cmdPlu(index).Picture <> App.Path & "\icons\upArr.bmp" Then
        cmdTot(0).Enabled = True
        cmdTot(1).Enabled = True
        ActiveReadServer "Select isnull(Whole_Unit,0) as Whole_Unit,Weight_Empty,Weight_Full from Products where Product_Code = '" & cmdPlu(index).Tag & "'"
        Whole_Unit = rs.Fields("Whole_Unit")
        lblWeight_Full.Caption = Val(rs.Fields("Weight_Full") & "")
        lblWeight_Empty.Caption = Val(rs.Fields("Weight_Empty") & "")
        rs.Close
        For i = 0 To cmdPlu.Count - 1
            If index <> i Then
                cmdPlu(i).Value = 0
            End If
        Next i
        Select Case Mid(cmdPlu(index).Caption, InStrRev(cmdPlu(index).Caption, " ") + 1)
            Case "500ml", "750ml", "1000ml"
                If Val(lblWeight_Full.Caption) <> 0 And Val(lblWeight_Empty.Caption) <> 0 Then
                    cmdTot(0).Caption = "Weight"
                Else
                    cmdTot(0).Caption = "Tots"
                End If
                cmdTot(1).Caption = Mid(cmdPlu(index).Caption, InStrRev(cmdPlu(index).Caption, " ") + 1)
            Case "each"
                cmdTot(0).Caption = "each"
                cmdTot(1).Enabled = False
                cmdTot(1).Caption = "each"
            Case Else
                If InStr(Mid(cmdPlu(index).Caption, InStrRev(cmdPlu(index).Caption, " ") + 1), "ml") > 0 Then
                    If Whole_Unit = 1 Then cmdTot(1).Enabled = False
                    cmdTot(0).Caption = "each"
                    cmdTot(1).Caption = "ml"
                    Exit Sub
                End If
                If InStr(Mid(cmdPlu(index).Caption, InStrRev(cmdPlu(index).Caption, " ") + 1), "lt") > 0 Then
                    If Whole_Unit = 1 Then cmdTot(1).Enabled = False
                    cmdTot(0).Caption = "each"
                    cmdTot(1).Caption = "lt"
                    Exit Sub
                End If
                If InStr(Mid(cmdPlu(index).Caption, InStrRev(cmdPlu(index).Caption, " ") + 1), "kg") > 0 Then
                    If Whole_Unit = 1 Then cmdTot(1).Enabled = False
                    cmdTot(0).Caption = "each"
                    cmdTot(1).Caption = "kg"
                    Exit Sub
                End If
                If InStr(Mid(cmdPlu(index).Caption, InStrRev(cmdPlu(index).Caption, " ") + 1), "g") > 0 Then
                    If Whole_Unit = 1 Then cmdTot(1).Enabled = False
                    cmdTot(0).Caption = "each"
                    cmdTot(1).Caption = "g"
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub cmdPlu_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdPlu(index).Picture <> App.Path & "\icons\downArr.bmp" And cmdPlu(index).Picture <> App.Path & "\icons\upArr.bmp" Then
        If cmdTot(0).Caption = "" Then cmdTot(0).Enabled = False
        If cmdTot(1).Caption = "" Then cmdTot(1).Enabled = False
    End If
    lblDetail.Tag = index
    DoEvents
    If cmdPlu(index).Value = 0 Then
        cmdTot(0).Enabled = False
        cmdTot(1).Enabled = False
    End If
End Sub
Private Sub cmdTot_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Right(lblDetail.Caption, 3) = " " & Chr(215) & " " Then
        lblDetail.Caption = Left(lblDetail.Caption, Len(lblDetail.Caption) - 3)
    End If
    DoEvents
    lblDetail.Font.Size = 16
    cmdTot(0).Enabled = False
    cmdTot(1).Enabled = False
    For i = 0 To cmdPlu.Count - 1
       cmdPlu(i).Value = 0
    Next i
    If Asc(Mid(lblDetail.Caption, 1, 1)) < 48 Or Asc(Mid(lblDetail.Caption, 1, 1)) > 57 Then
        If InStr(lblDetail.Caption, ".") = 0 Then
            lblDetail.Caption = "1"
        End If
    End If
    Take_No = 0
    Qty_Counted = lblDetail.Caption
    Select Case cmdTot(index).Caption
        Case "Weight"
            Unit_of_Measure = "Tots"
            BContents = Val(lblWeight_Full.Caption) - Val(lblWeight_Empty.Caption)
            ActiveReadServer "Select Unit_Size from Products where Product_code = '" & cmdPlu(Val(lblDetail.Tag)).Tag & "'"
            If rs.RecordCount > 0 Then
                If Val(lblDetail.Caption) <> 0 Then
                    TotWeight = BContents / (Val(rs.Fields("Unit_Size") & "") / 25)
                    Tots = Round(((lblDetail.Caption - Val(lblWeight_Empty.Caption)) / TotWeight))
                    Qty_Counted = Tots
                Else
                    Tots = 0
                    Qty_Counted = Tots
                End If
            End If
            rs.Close
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Val(lblDetail.Tag)).Caption
            grdMain.TextMatrix(grdMain.Row, 1) = Tots & " (Tots)"
            grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Val(lblDetail.Tag)).Tag
            lblDetail.Caption = "Counted: " & Tots & " x " & cmdPlu(Val(lblDetail.Tag)).Caption & " (Tots)"
        Case "each"
            Unit_of_Measure = "each"
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Val(lblDetail.Tag)).Caption
            grdMain.TextMatrix(grdMain.Row, 1) = lblDetail.Caption
            grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Val(lblDetail.Tag)).Tag
            lblDetail.Caption = "Counted: " & lblDetail.Caption & " x " & cmdPlu(Val(lblDetail.Tag)).Caption
        Case "ml"
            Unit_of_Measure = "ml"
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Val(lblDetail.Tag)).Caption
            grdMain.TextMatrix(grdMain.Row, 1) = lblDetail.Caption & "ml"
            grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Val(lblDetail.Tag)).Tag
            lblDetail.Caption = "Counted: " & lblDetail.Caption & " x " & cmdPlu(Val(lblDetail.Tag)).Caption & "ml"
        Case "Cases"
            Unit_of_Measure = "Cases"
        Case "500ml", "750ml", "1000ml"
            Unit_of_Measure = ""
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Val(lblDetail.Tag)).Caption
            grdMain.TextMatrix(grdMain.Row, 1) = lblDetail.Caption
            grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Val(lblDetail.Tag)).Tag
            lblDetail.Caption = "Counted: " & lblDetail.Caption & " x " & cmdPlu(Val(lblDetail.Tag)).Caption
        Case "Tots"
            Unit_of_Measure = "Tots"
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Val(lblDetail.Tag)).Caption
            grdMain.TextMatrix(grdMain.Row, 1) = lblDetail.Caption & " (Tots)"
            grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Val(lblDetail.Tag)).Tag
            lblDetail.Caption = "Counted: " & lblDetail.Caption & " x " & cmdPlu(Val(lblDetail.Tag)).Caption & " (Tots)"
        Case "g"
            Unit_of_Measure = "g"
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Val(lblDetail.Tag)).Caption
            grdMain.TextMatrix(grdMain.Row, 1) = lblDetail.Caption & "g"
            grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Val(lblDetail.Tag)).Tag
            lblDetail.Caption = "Counted: " & lblDetail.Caption & " x " & cmdPlu(Val(lblDetail.Tag)).Caption
        Case "kg"
            Unit_of_Measure = "kg"
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Val(lblDetail.Tag)).Caption
            grdMain.TextMatrix(grdMain.Row, 1) = lblDetail.Caption & "kg"
            grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Val(lblDetail.Tag)).Tag
            lblDetail.Caption = "Counted: " & lblDetail.Caption & " x " & cmdPlu(Val(lblDetail.Tag)).Caption
        Case "lt"
            Unit_of_Measure = "lt"
            grdMain.Rows = grdMain.Rows + 1
            grdMain.Row = grdMain.Rows - 1
            grdMain.TextMatrix(grdMain.Row, 0) = cmdPlu(Val(lblDetail.Tag)).Caption
            grdMain.TextMatrix(grdMain.Row, 1) = lblDetail.Caption & "lt"
            grdMain.TextMatrix(grdMain.Row, 2) = cmdPlu(Val(lblDetail.Tag)).Tag
            lblDetail.Caption = "Counted: " & lblDetail.Caption & " x " & cmdPlu(Val(lblDetail.Tag)).Caption
    End Select
    cmdPlu(Val(lblDetail.Tag)).TextDescrCB.Text = ""
    ActiveUpdateServer "INSERT INTO Stock_Take_Listing (Date_Time,  Workstation_No, User_No, Location_No,Product_Code,Department_No,Qty_Counted,Unit_of_Measure) values " & _
    "(Getdate()," & Workstation_No & "," & UserRecord.User_Number & "," & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & ",'" & cmdPlu(Val(lblDetail.Tag)).Tag & "','" & Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) & "'," & Qty_Counted & ",'" & Unit_of_Measure & "')"
    
    ActiveReadServer "Select max(Line_No) as Line_No from Stock_Take_Listing"
    grdMain.TextMatrix(grdMain.Row, 3) = rs.Fields("Line_No")
    rs.Close
    ActiveReadServer "Select * from Departments_Stock where Location_No = " & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1) & " and Department_No = '" & Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) & "'"
    If rs.RecordCount > 0 Then
        cmdDeptStrip(Val(lblDepart.Tag)).TextDescrCB.Text = rs.Fields("Product_Count") & " Counted of " & rs.Fields("Product_List")
        For i = 0 To grdDept.Rows - 1
            If grdDept.TextMatrix(i, 1) = Mid(lblDepart.Caption, 11, InStr(lblDepart, ">") - 12) Then
                grdDept.TextMatrix(i, 3) = rs.Fields("Product_Count") & " Counted of " & rs.Fields("Product_List")
            End If
        Next i
    End If
    rs.Close
    grdMain.ShowCell grdMain.Rows - 1, 0
    Me.SetFocus
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
            ActiveReadServer "Select *,isnull((Select count(Location_No) from Stock_Take_Listing where Stock_Take_Listing.Location_No = Locations.Location_No group by Location_No),0)as Products from Locations order by Location_No"
            While Not rs.EOF
                grdLoc.Rows = grdLoc.Rows + 1
                grdLoc.TextMatrix(grdLoc.Rows - 1, 0) = rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                Select Case rs.Fields("Products")
                    Case 0: grdLoc.TextMatrix(grdLoc.Rows - 1, 1) = "No Active Count"
                    Case Else: grdLoc.TextMatrix(grdLoc.Rows - 1, 1) = "Active Count"
                End Select
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
        grdMerge.FontSize = 10
    End If
    grdMerge.ColWidth(0) = grdMerge.Width * 0.4
    grdMerge.ColWidth(1) = grdMerge.Width * 0.2
    grdMerge.ColWidth(2) = grdMerge.Width * 0.2
    grdMerge.ColWidth(3) = grdMerge.Width * 0.2
    grdMain.ColWidth(0) = grdMain.Width * 0.7
    grdMain.ColWidth(1) = grdMain.Width * 0.3
    grdRem.ColWidth(0) = grdRem.Width * 0.7
    grdRem.ColWidth(1) = grdRem.Width * 0.3
    
    QAnswer.Caption = ""
    lblDetail.Caption = "Select a Location to Count"
    For i = 0 To cmdDeptStrip.Count - 1
        cmdDeptStrip(i).Visible = False
    Next i
    cmdUser.Caption = "<All Locations>"
    grdMain.Rows = 1
    For i = 0 To cmdPlu.Count - 1
        cmdPlu(i).Visible = False
    Next i
    cmdTot(0).Enabled = False
    cmdTot(1).Enabled = False
    picMerge.Visible = False
    grdMerge.Rows = 1
End Sub
Private Sub cmdArrow_MouseUp(index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    scrolTimer.Enabled = False
    scrolTimer.Interval = 700
End Sub
Private Sub Form_Load()
    lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
    lblDetail.Caption = "Select a Location to Count"
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
    grdMerge.TextMatrix(0, 0) = "Description"
    grdMerge.TextMatrix(0, 1) = "On Hand"
    grdMerge.TextMatrix(0, 2) = "Counted"
    grdMerge.TextMatrix(0, 3) = "Variance"
    grdMerge.TextMatrix(0, 4) = "Product"
    grdMerge.ColHidden(4) = True
    grdMerge.ColHidden(5) = True
    grdMerge.ColHidden(6) = True
    grdMerge.ColHidden(7) = True
    grdMerge.ColAlignment(0) = flexAlignLeftCenter
    grdMerge.ColAlignment(1) = flexAlignRightCenter
    grdMerge.ColAlignment(2) = flexAlignRightCenter
    grdMerge.ColAlignment(3) = flexAlignRightCenter
End Sub
Private Sub grdLoc_Click()
    cmdUser.Caption = grdLoc.TextMatrix(grdLoc.Row, 0)
    lblDetail.Caption = grdLoc.TextMatrix(grdLoc.Row, 1)
    grdLoc.Visible = False
    For i = 0 To cmdPlu.Count - 1
        cmdPlu(i).Visible = False
    Next i
    cmdTot(0).Enabled = False
    cmdTot(1).Enabled = False
    lblDepart.Tag = ""
    picMerge.Visible = False
    grdMerge.Rows = 1
    Load_Departments Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
    Load_List 0, Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
    ActiveReadServer "Select * from Location_Stock where Location_No=" & Mid(grdLoc.TextMatrix(grdLoc.Row, 0), 1, InStr(grdLoc.TextMatrix(grdLoc.Row, 0), "-") - 1)
    If rs.RecordCount > 0 Then
        lblDepart.Caption = cmdUser.Caption & " > " & rs.Fields("Product_Count") & " Counted of " & rs.Fields("Product_List")
    End If
    rs.Close
End Sub
Private Sub Load_List(Department_No, Location_No)
    Screen.MousePointer = 11
    Select Case Department_No
        Case 0
            ActiveReadServer "Select * from Stock_Take_Listing_View where Location_No=" & Location_No & " order by Line_No"
        Case Else
            ActiveReadServer "Select * from Stock_Take_Listing_View where Location_No=" & Location_No & " and Department_No='" & Department_No & "' order by Line_No"
    End Select
    grdMain.Rows = 1
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.Row = grdMain.Rows - 1
        Select Case rs.Fields("Unit_Size")
            Case 0
                unit_Size = rs.Fields("Unit_of_Measure")
            Case Else
                unit_Size = rs.Fields("Unit_Size") & rs.Fields("U_O_M")
        End Select
        grdMain.TextMatrix(grdMain.Row, 0) = rs.Fields("Description") & " " & unit_Size
        Select Case rs.Fields("Unit_of_Measure") & ""
            Case "", "each"
                grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Qty_Counted")
            Case "Tots"
                grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Qty_Counted") & " (Tots)"
            Case "Cases"
                grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Qty_Counted") & " (Cases)"
            Case Else
                grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Qty_Counted") & rs.Fields("Unit_of_Measure") & ""
        End Select
        grdMain.TextMatrix(grdMain.Row, 2) = rs.Fields("Product_Code")
        For i = 0 To cmdPlu.Count - 1
            If rs.Fields("Product_Code") = cmdPlu(i).Tag Then
                'cmdPlu(i).TextDescrCB.OffsetY = -6
                'cmdPlu(i).TextDescrCB.Text = "Counted"
                'cmdPlu(i).TextDescrCB.ColorNormal = vbBlue
            End If
        Next i
        grdMain.TextMatrix(grdMain.Row, 3) = rs.Fields("Line_No")
        rs.MoveNext
    Wend
    rs.Close
    grdMain.ShowCell grdMain.Rows - 1, 0
    Screen.MousePointer = 0
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
            cmdDeptStrip(i).TextDescrCB.Text = rs.Fields("Product_Count") & " Counted of " & rs.Fields("Product_List")
            cmdDeptStrip(i).TextDescrCB.ColorNormal = vbYellow
            cmdDeptStrip(i).TextDescrCB.OffsetY = -6
            If cmdDeptStrip(i).Visible = False Then cmdDeptStrip(i).Visible = True
            grdDept.Row = grdDept.Rows - 1
            grdDept.TextMatrix(grdDept.Rows - 1, 0) = UCase(Replace(rs.Fields("Short_Name"), "&", "&&"))
            grdDept.TextMatrix(grdDept.Rows - 1, 1) = rs.Fields("Department_No")
            grdDept.TextMatrix(grdDept.Rows - 1, 2) = rs.Fields("Dept_Parent")
            grdDept.TextMatrix(grdDept.Rows - 1, 3) = rs.Fields("Product_Count") & " Counted of " & rs.Fields("Product_List")
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
            grdDept.TextMatrix(grdDept.Rows - 1, 3) = rs.Fields("Product_Count") & " Counted of " & rs.Fields("Product_List")
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

Private Sub ScrolTimer1_Timer()
    scrolTimer1.Interval = 50
    Select Case scrolTimer1.Tag
        Case "0"
            If grdMerge.Row <> 1 Then
                grdMerge.Row = grdMerge.Row - 1
            End If
        Case "1"
            If grdMerge.Row <> grdMerge.Rows - 1 Then
                grdMerge.Row = grdMerge.Row + 1
            End If
    End Select
    grdMerge.ShowCell grdMerge.Row, 0
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
