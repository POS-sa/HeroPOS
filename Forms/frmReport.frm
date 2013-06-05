VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmReports 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11880
   ScaleWidth      =   15780
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicTradecomparison 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8265
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   14025
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   14055
      Begin VB.TextBox txtcount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   30
         TabIndex        =   50
         Top             =   6840
         Width           =   2235
      End
      Begin VSFlex8Ctl.VSFlexGrid GrdCompare 
         Height          =   1455
         Left            =   30
         TabIndex        =   45
         Top             =   5400
         Width           =   13545
         _cx             =   23892
         _cy             =   2566
         Appearance      =   0
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid GrdAvespend 
         Height          =   1665
         Left            =   30
         TabIndex        =   47
         Top             =   3750
         Width           =   13545
         _cx             =   23892
         _cy             =   2937
         Appearance      =   0
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid Grdcovers 
         Height          =   1245
         Left            =   30
         TabIndex        =   49
         Top             =   2520
         Width           =   13545
         _cx             =   23892
         _cy             =   2196
         Appearance      =   0
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid GrdTodayTotal 
         Height          =   285
         Left            =   30
         TabIndex        =   46
         Top             =   2280
         Width           =   13545
         _cx             =   23892
         _cy             =   503
         Appearance      =   0
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid GrdToday 
         Height          =   1725
         Left            =   30
         TabIndex        =   48
         Top             =   570
         Width           =   13545
         _cx             =   23892
         _cy             =   3043
         Appearance      =   0
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
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VB.Label LblCaption2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sales Reports."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   30
         TabIndex        =   51
         Top             =   240
         Width           =   13515
      End
   End
   Begin VB.PictureBox Pictimeback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   9870
      ScaleHeight     =   1125
      ScaleWidth      =   3585
      TabIndex        =   40
      Top             =   1080
      Width           =   3585
      Begin MSComCtl2.DTPicker TStart 
         Height          =   375
         Left            =   210
         TabIndex        =   41
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy (ddd)"
         Format          =   66453506
         CurrentDate     =   39414
      End
      Begin btButtonEx.ButtonEx Boktime 
         Height          =   330
         Left            =   2220
         TabIndex        =   42
         Top             =   630
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Appearance      =   3
         Caption         =   "Ok"
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
      Begin MSComCtl2.DTPicker TEnd 
         Height          =   375
         Left            =   1860
         TabIndex        =   43
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy (ddd)"
         Format          =   66453506
         CurrentDate     =   39414
      End
      Begin MSForms.Image Image8 
         Height          =   945
         Left            =   60
         Top             =   60
         Width           =   3435
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "6059;1667"
      End
      Begin MSForms.Image Pictime 
         Height          =   1065
         Left            =   0
         Top             =   0
         Width           =   3555
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "6271;1879"
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   180
      Top             =   9630
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   375
      Left            =   6900
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   660
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   8421504
      Caption         =   "¦"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VB.PictureBox picBlocDate 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   4500
      ScaleHeight     =   435
      ScaleWidth      =   2805
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   4500
      ScaleHeight     =   1065
      ScaleWidth      =   5355
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   5355
      Begin MSComCtl2.DTPicker mthViewStart 
         Height          =   375
         Left            =   180
         TabIndex        =   35
         Top             =   150
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy (ddd)"
         Format          =   66453507
         CurrentDate     =   39414
      End
      Begin btButtonEx.ButtonEx cmdOk 
         Height          =   330
         Left            =   4080
         TabIndex        =   17
         Top             =   570
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Appearance      =   3
         Caption         =   "Ok"
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
      Begin MSComCtl2.DTPicker mthViewEnd 
         Height          =   375
         Left            =   2730
         TabIndex        =   36
         Top             =   150
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy (ddd)"
         Format          =   66453507
         CurrentDate     =   39414
      End
      Begin MSForms.Image Image6 
         Height          =   945
         Left            =   60
         Top             =   60
         Width           =   5235
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "9234;1667"
      End
      Begin MSForms.Image Image5 
         Height          =   1065
         Left            =   0
         Top             =   0
         Width           =   5355
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "9446;1879"
      End
   End
   Begin VB.PictureBox picLive 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDE6DF&
      ForeColor       =   &H80000008&
      Height          =   385
      Left            =   2370
      ScaleHeight     =   360
      ScaleWidth      =   4125
      TabIndex        =   30
      Top             =   1210
      Visible         =   0   'False
      Width           =   4160
      Begin VB.Label lblLive 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You have no Open Tables"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   31
         Top             =   50
         Width           =   4035
      End
   End
   Begin BTNENHLib4.BtnEnh cmdMenuSep 
      Height          =   465
      Left            =   6480
      TabIndex        =   6
      Top             =   0
      Width           =   135
      _Version        =   524298
      _ExtentX        =   238
      _ExtentY        =   820
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
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
      Surface         =   10
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":0000
      textLT          =   "frmReport.frx":0018
      textCT          =   "frmReport.frx":0030
      textRT          =   "frmReport.frx":0048
      textLM          =   "frmReport.frx":0060
      textRM          =   "frmReport.frx":0078
      textLB          =   "frmReport.frx":0090
      textCB          =   "frmReport.frx":00A8
      textRB          =   "frmReport.frx":00C0
      colorBack       =   "frmReport.frx":00D8
      colorIntern     =   "frmReport.frx":0102
      colorMO         =   "frmReport.frx":012C
      colorFocus      =   "frmReport.frx":0156
      colorDisabled   =   "frmReport.frx":0180
      colorPressed    =   "frmReport.frx":01AA
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdMenu 
      Height          =   465
      Index           =   4
      Left            =   6630
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1665
      _Version        =   524298
      _ExtentX        =   2937
      _ExtentY        =   820
      _StockProps     =   66
      Caption         =   "Time Sheets"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
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
      Surface         =   10
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":01D4
      textLT          =   "frmReport.frx":024A
      textCT          =   "frmReport.frx":0262
      textRT          =   "frmReport.frx":027A
      textLM          =   "frmReport.frx":0292
      textRM          =   "frmReport.frx":02AA
      textLB          =   "frmReport.frx":02C2
      textCB          =   "frmReport.frx":02DA
      textRB          =   "frmReport.frx":02F2
      colorBack       =   "frmReport.frx":030A
      colorIntern     =   "frmReport.frx":0334
      colorMO         =   "frmReport.frx":035E
      colorFocus      =   "frmReport.frx":0388
      colorDisabled   =   "frmReport.frx":03B2
      colorPressed    =   "frmReport.frx":03DC
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdMenu 
      Height          =   465
      Index           =   5
      Left            =   8310
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1995
      _Version        =   524298
      _ExtentX        =   3519
      _ExtentY        =   820
      _StockProps     =   66
      Caption         =   "Journals"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
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
      Surface         =   10
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":0406
      textLT          =   "frmReport.frx":0476
      textCT          =   "frmReport.frx":048E
      textRT          =   "frmReport.frx":04A6
      textLM          =   "frmReport.frx":04BE
      textRM          =   "frmReport.frx":04D6
      textLB          =   "frmReport.frx":04EE
      textCB          =   "frmReport.frx":0506
      textRB          =   "frmReport.frx":051E
      colorBack       =   "frmReport.frx":0536
      colorIntern     =   "frmReport.frx":0560
      colorMO         =   "frmReport.frx":058A
      colorFocus      =   "frmReport.frx":05B4
      colorDisabled   =   "frmReport.frx":05DE
      colorPressed    =   "frmReport.frx":0608
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdMenu 
      Height          =   465
      Index           =   6
      Left            =   10320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   820
      _StockProps     =   66
      Caption         =   "Debtors"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
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
      Surface         =   10
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":0632
      textLT          =   "frmReport.frx":06A0
      textCT          =   "frmReport.frx":06B8
      textRT          =   "frmReport.frx":06D0
      textLM          =   "frmReport.frx":06E8
      textRM          =   "frmReport.frx":0700
      textLB          =   "frmReport.frx":0718
      textCB          =   "frmReport.frx":0730
      textRB          =   "frmReport.frx":0748
      colorBack       =   "frmReport.frx":0760
      colorIntern     =   "frmReport.frx":078A
      colorMO         =   "frmReport.frx":07B4
      colorFocus      =   "frmReport.frx":07DE
      colorDisabled   =   "frmReport.frx":0808
      colorPressed    =   "frmReport.frx":0832
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdMenu 
      Height          =   465
      Index           =   7
      Left            =   11940
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1785
      _Version        =   524298
      _ExtentX        =   3149
      _ExtentY        =   820
      _StockProps     =   66
      Caption         =   "Stock Takes"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
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
      Surface         =   10
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":085C
      textLT          =   "frmReport.frx":08D2
      textCT          =   "frmReport.frx":08EA
      textRT          =   "frmReport.frx":0902
      textLM          =   "frmReport.frx":091A
      textRM          =   "frmReport.frx":0932
      textLB          =   "frmReport.frx":094A
      textCB          =   "frmReport.frx":0962
      textRB          =   "frmReport.frx":097A
      colorBack       =   "frmReport.frx":0992
      colorIntern     =   "frmReport.frx":09BC
      colorMO         =   "frmReport.frx":09E6
      colorFocus      =   "frmReport.frx":0A10
      colorDisabled   =   "frmReport.frx":0A3A
      colorPressed    =   "frmReport.frx":0A64
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdMenu 
      Height          =   465
      Index           =   3
      Left            =   4830
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1635
      _Version        =   524298
      _ExtentX        =   2884
      _ExtentY        =   820
      _StockProps     =   66
      Caption         =   "Rooms"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
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
      Surface         =   10
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":0A8E
      textLT          =   "frmReport.frx":0AF8
      textCT          =   "frmReport.frx":0B10
      textRT          =   "frmReport.frx":0B28
      textLM          =   "frmReport.frx":0B40
      textRM          =   "frmReport.frx":0B58
      textLB          =   "frmReport.frx":0B70
      textCB          =   "frmReport.frx":0B88
      textRB          =   "frmReport.frx":0BA0
      colorBack       =   "frmReport.frx":0BB8
      colorIntern     =   "frmReport.frx":0BE2
      colorMO         =   "frmReport.frx":0C0C
      colorFocus      =   "frmReport.frx":0C36
      colorDisabled   =   "frmReport.frx":0C60
      colorPressed    =   "frmReport.frx":0C8A
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdMenu 
      Height          =   465
      Index           =   2
      Left            =   3210
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   820
      _StockProps     =   66
      Caption         =   "Purchases"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
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
      Surface         =   10
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":0CB4
      textLT          =   "frmReport.frx":0D26
      textCT          =   "frmReport.frx":0D3E
      textRT          =   "frmReport.frx":0D56
      textLM          =   "frmReport.frx":0D6E
      textRM          =   "frmReport.frx":0D86
      textLB          =   "frmReport.frx":0D9E
      textCB          =   "frmReport.frx":0DB6
      textRB          =   "frmReport.frx":0DCE
      colorBack       =   "frmReport.frx":0DE6
      colorIntern     =   "frmReport.frx":0E10
      colorMO         =   "frmReport.frx":0E3A
      colorFocus      =   "frmReport.frx":0E64
      colorDisabled   =   "frmReport.frx":0E8E
      colorPressed    =   "frmReport.frx":0EB8
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdMenu 
      Height          =   465
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1575
      _Version        =   524298
      _ExtentX        =   2778
      _ExtentY        =   820
      _StockProps     =   66
      Caption         =   "Stock"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
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
      Surface         =   10
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":0EE2
      textLT          =   "frmReport.frx":0F4C
      textCT          =   "frmReport.frx":0F64
      textRT          =   "frmReport.frx":0F7C
      textLM          =   "frmReport.frx":0F94
      textRM          =   "frmReport.frx":0FAC
      textLB          =   "frmReport.frx":0FC4
      textCB          =   "frmReport.frx":0FDC
      textRB          =   "frmReport.frx":0FF4
      colorBack       =   "frmReport.frx":100C
      colorIntern     =   "frmReport.frx":1036
      colorMO         =   "frmReport.frx":1060
      colorFocus      =   "frmReport.frx":108A
      colorDisabled   =   "frmReport.frx":10B4
      colorPressed    =   "frmReport.frx":10DE
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
   End
   Begin BTNENHLib4.BtnEnh cmdMenu 
      Height          =   465
      Index           =   0
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   820
      _StockProps     =   66
      Caption         =   "Sales"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      Surface         =   10
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmReport.frx":1108
      textLT          =   "frmReport.frx":1172
      textCT          =   "frmReport.frx":118A
      textRT          =   "frmReport.frx":11A2
      textLM          =   "frmReport.frx":11BA
      textRM          =   "frmReport.frx":11D2
      textLB          =   "frmReport.frx":11EA
      textCB          =   "frmReport.frx":1202
      textRB          =   "frmReport.frx":121A
      colorBack       =   "frmReport.frx":1232
      colorIntern     =   "frmReport.frx":125C
      colorMO         =   "frmReport.frx":1286
      colorFocus      =   "frmReport.frx":12B0
      colorDisabled   =   "frmReport.frx":12DA
      colorPressed    =   "frmReport.frx":1304
      Style           =   2
      Orientation     =   2
      HollowFrame     =   -1  'True
      LightDirection  =   1
      RectHardEdges   =   -1  'True
      Value           =   -1  'True
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8220
      Left            =   30
      ScaleHeight     =   8220
      ScaleWidth      =   13755
      TabIndex        =   18
      Top             =   1170
      Visible         =   0   'False
      Width           =   13755
      Begin VSFlex8Ctl.VSFlexGrid grdMain 
         Height          =   7710
         Left            =   -40
         TabIndex        =   19
         Top             =   0
         Width           =   13695
         _cx             =   24156
         _cy             =   13600
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   12632256
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
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":132E
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
      Begin VSFlex8Ctl.VSFlexGrid grdTotal 
         Height          =   705
         Left            =   0
         TabIndex        =   20
         Top             =   7710
         Width           =   13695
         _cx             =   24156
         _cy             =   1244
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   430
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":13A6
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
         Begin btButtonEx.ButtonEx cmdSearch 
            Height          =   375
            Left            =   8220
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Appearance      =   3
            BorderColor     =   4210752
            Caption         =   "Search"
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
         Begin btButtonEx.ButtonEx cmdPrice 
            Height          =   375
            Left            =   6510
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   661
            Appearance      =   3
            BorderColor     =   4210752
            Caption         =   "Change Selling Price"
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
      End
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8220
      Left            =   0
      ScaleHeight     =   8190
      ScaleWidth      =   13665
      TabIndex        =   21
      Top             =   1140
      Visible         =   0   'False
      Width           =   13695
      Begin VSFlex8Ctl.VSFlexGrid grdRev 
         Height          =   2655
         Left            =   60
         TabIndex        =   22
         Top             =   60
         Width           =   6450
         _cx             =   11377
         _cy             =   4683
         Appearance      =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   8421504
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   2300
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":141E
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid grdTrans 
         Height          =   2265
         Left            =   60
         TabIndex        =   23
         Top             =   2790
         Width           =   6450
         _cx             =   11377
         _cy             =   3995
         Appearance      =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   8421504
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   2300
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":1472
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid grdCount 
         Height          =   3015
         Left            =   60
         TabIndex        =   24
         Top             =   5130
         Width           =   6450
         _cx             =   11377
         _cy             =   5318
         Appearance      =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   8421504
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   8
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   2300
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":14C6
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid grdTax 
         Height          =   2265
         Left            =   6570
         TabIndex        =   25
         Top             =   60
         Width           =   7065
         _cx             =   12462
         _cy             =   3995
         Appearance      =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   8421504
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   5000
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":151A
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid grdStock 
         Height          =   2655
         Left            =   6570
         TabIndex        =   26
         Top             =   2400
         Width           =   7050
         _cx             =   12435
         _cy             =   4683
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   8421504
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   5000
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":156E
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid grdCred 
         Height          =   2265
         Left            =   6570
         TabIndex        =   27
         Top             =   5130
         Width           =   7065
         _cx             =   12462
         _cy             =   3995
         Appearance      =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   8421504
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   1750
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":15C2
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VSFlex8Ctl.VSFlexGrid grdGP 
         Height          =   705
         Left            =   6570
         TabIndex        =   28
         Top             =   7440
         Width           =   7065
         _cx             =   12462
         _cy             =   1244
         Appearance      =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   15329975
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   8421504
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   340
         RowHeightMax    =   0
         ColWidthMin     =   1750
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmReport.frx":1616
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   2
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
   End
   Begin VB.PictureBox picAnalysis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8225
      Left            =   0
      ScaleHeight     =   8190
      ScaleWidth      =   13695
      TabIndex        =   15
      Top             =   1140
      Visible         =   0   'False
      Width           =   13725
      Begin VB.Line Line2 
         X1              =   6960
         X2              =   6960
         Y1              =   4680
         Y2              =   8190
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   13710
         Y1              =   4680
         Y2              =   4680
      End
   End
   Begin btButtonEx.ButtonEx Buttime 
      Height          =   375
      Left            =   870
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   9660
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   8421504
      Caption         =   "¦"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VB.Label Lbltime 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   39
      Top             =   10710
      Width           =   2085
   End
   Begin MSForms.Label Lbltimeback 
      Height          =   375
      Left            =   30
      TabIndex        =   38
      Top             =   10230
      Width           =   2235
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      PicturePosition =   131072
      Size            =   "3942;661"
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblDate 
      Height          =   315
      Left            =   4560
      TabIndex        =   13
      Top             =   750
      Width           =   2235
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "1 Feb 2006 to 13 Feb 2006"
      Size            =   "3942;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Report_No 
      Caption         =   "0"
      Height          =   345
      Left            =   2430
      TabIndex        =   29
      Top             =   570
      Visible         =   0   'False
      Width           =   1965
   End
   Begin MSForms.ComboBox cmb3 
      Height          =   375
      Left            =   11730
      TabIndex        =   14
      Top             =   660
      Width           =   1905
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3360;661"
      ListWidth       =   4938
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image2 
      Height          =   375
      Left            =   4500
      Top             =   660
      Width           =   2385
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "4207;661"
   End
   Begin MSForms.ComboBox cmb1 
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   660
      Width           =   1845
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3254;661"
      ListRows        =   10
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmb2 
      Height          =   375
      Left            =   7380
      TabIndex        =   0
      Top             =   660
      Width           =   2415
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4260;661"
      ListRows        =   10
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      Caption         =   "Sales Reports."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   4395
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   2
      Left            =   4140
      Top             =   990
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   1
      Left            =   3810
      Top             =   990
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   0
      Left            =   3480
      Top             =   990
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image3 
      Height          =   90
      Left            =   60
      Top             =   990
      Width           =   3375
      BackColor       =   16761024
      Size            =   "5953;159"
   End
   Begin MSForms.Image Image1 
      Height          =   705
      Left            =   -60
      Top             =   510
      Width           =   13785
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "24315;1244"
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Stock_on_Hand_Suppliers()
    If cmb1.Text <> "<All Suppliers>" Then
        LocString = Trim(Mid(cmb1.Text, InStr(cmb1.Text, "-") + 1))
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    grdTotal.TextMatrix(0, 0) = ""
    grdTotal.TextMatrix(0, 1) = ""
    grdTotal.TextMatrix(0, 2) = ""
    grdTotal.TextMatrix(0, 3) = ""
    grdTotal.TextMatrix(0, 4) = "0.000"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    grdTotal.TextMatrix(0, 8) = "0.00"
    ActiveReadServer "Select * from SOH_Supplier_View where " & _
    " Department_No like '" & DeptString & "'" & _
    " and Supplier_No like '" & LocString & "' order by Description"
    grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 1) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Products = " & rs.RecordCount
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Unit_Size") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("Supplier") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(rs.Fields("Stock_on_Hand") & "", 3)
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Ave_Cost") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Total_Excl") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = rs.Fields("Sales_Tax") & "%"
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = Format(rs.Fields("Total_Incl") & "", "0.00")
        grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), 3)
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), "0.00")
        grdTotal.TextMatrix(0, 8) = Format(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub
Private Sub Pack_Links_Screen()
    picBlocDate.Visible = True
    picMain.Visible = True
    picMain.Visible = True
    cmb2.Width = 3225
    cmb1.Width = 2925
    cmb2.Left = 7380
    cmb1.Left = 10680
    cmb3.Visible = False
    grdMain.Cols = 7
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Product Code"
    grdMain.TextMatrix(0, 1) = " Description"
    grdMain.TextMatrix(0, 2) = " Department"
    grdMain.TextMatrix(0, 3) = " Pack Size"
    grdMain.TextMatrix(0, 4) = " Ave Cost"
    grdMain.TextMatrix(0, 5) = " Link Code"
    grdMain.TextMatrix(0, 6) = " Link Description"
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignLeftCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignRightCenter
    grdMain.ColAlignment(5) = flexAlignLeftCenter
    grdMain.ColAlignment(6) = flexAlignLeftCenter
    grdMain.ColWidth(0) = grdMain.Width * 0.1
    grdMain.ColWidth(1) = grdMain.Width * 0.22
    grdMain.ColWidth(2) = grdMain.Width * 0.2
    grdMain.ColWidth(3) = grdMain.Width * 0.08
    grdMain.ColWidth(4) = grdMain.Width * 0.08
    grdMain.ColWidth(5) = grdMain.Width * 0.1
    grdMain.ColWidth(6) = grdMain.Width * 0.22
    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    lblCaption.Caption = "Reports - Pack Links"
    
    cmb1.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    cmb1.AddItem "<All Locations>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<All Locations>"
    
    cmb3.Clear
    ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
    cmb3.AddItem "<All Departments>"
    While Not rs.EOF
        cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb3.Text = "<All Departments>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Stock_on_Hand()
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    grdTotal.TextMatrix(0, 0) = ""
    grdTotal.TextMatrix(0, 1) = ""
    grdTotal.TextMatrix(0, 2) = ""
    grdTotal.TextMatrix(0, 3) = ""
    grdTotal.TextMatrix(0, 4) = "0.000"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    grdTotal.TextMatrix(0, 8) = "0.00"
    ActiveReadServer "Select * from SOH_View where " & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'"
    grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 1) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Products = " & rs.RecordCount
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Unit_Size") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("Department") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(rs.Fields("Stock_on_Hand") & "", 3)
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Ave_Cost") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Total_Excl") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = rs.Fields("Sales_Tax") & "%"
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = Format(rs.Fields("Total_Incl") & "", "0.00")
        grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), 3)
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), "0.00")
        grdTotal.TextMatrix(0, 8) = Format(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub

Private Sub Stock_Levels()
If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    grdTotal.TextMatrix(0, 0) = ""
    grdTotal.TextMatrix(0, 1) = ""
    grdTotal.TextMatrix(0, 2) = ""
    grdTotal.TextMatrix(0, 3) = ""
    grdTotal.TextMatrix(0, 4) = "0.000"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    grdTotal.TextMatrix(0, 8) = "0.00"
    ActiveReadServer "Select * from Stocklow_View where " & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'"
    grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 1) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Products = " & rs.RecordCount
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Unit_Size") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("Department_no") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(rs.Fields("Stock_on_Hand") & "", 3)
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Ave_Cost") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Total_Excl") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = rs.Fields("Sales_Tax") & "%"
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = Format(rs.Fields("Total_Incl") & "", "0.00")
        grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), 3)
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), "0.00")
        grdTotal.TextMatrix(0, 8) = Format(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub







Private Sub Product_Analysis()
    If cmb1.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2)
        End If
    End If
    ActiveReadServer "SELECT Products.Product_Code, Products.Description, Products.Landed_Cost, Products.Selling_Price, Departments.Department_No,Departments.Dept_Name, Products.Sales_Tax," & _
    " Case Products.Selling_Price WHEN 0 THEN 0 Else ((Products.Selling_Price / ((100 + Products.Sales_Tax) / 100)) - Products.Landed_Cost)/(Products.Selling_Price /((100 + Products.Sales_Tax)/100)) * 100" & _
    " END AS GP, Case Products.Landed_Cost WHEN 0 THEN 100 Else ((Products.Selling_Price / ((100 + Products.Sales_Tax) / 100)) - Products.Landed_Cost)/Products.Landed_Cost * 100" & _
    " END As Markup FROM Products LEFT OUTER JOIN Departments ON Products.Department_No = Departments.Department_No Where Sales_Item = 1 and Products.Department_No like '" & DeptString & "' order by Products.Department_No,Description"
    grdTotal.TextMatrix(0, 0) = "0"
    grdTotal.TextMatrix(0, 1) = "0"
    grdTotal.TextMatrix(0, 2) = "0"
    grdTotal.TextMatrix(0, 3) = "0"
    grdTotal.TextMatrix(0, 4) = "0"
    grdTotal.TextMatrix(0, 5) = "0"
    grdTotal.TextMatrix(0, 6) = "0"
    ToTGPValue = 0
    TotMarkupValue = 0
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Department_No") & " " & rs.Fields("Dept_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Landed_Cost"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(rs.Fields("Markup"), 2) & "%"
        TotMarkupValue = TotMarkupValue + rs.Fields("Markup")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Round(rs.Fields("GP"), 2) & "%"
        ToTGPValue = ToTGPValue + Round(rs.Fields("GP"), 2)
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Selling_Price"), "0.00")
        grdTotal.TextMatrix(0, 0) = rs.RecordCount - 1 & " Products"
        grdTotal.TextMatrix(0, 1) = rs.RecordCount - 1 & " Products"
        grdTotal.TextMatrix(0, 2) = rs.RecordCount - 1 & " Products"
        rs.MoveNext
    Wend
    grdTotal.TextMatrix(0, 3) = "Average Markup%: " & Round(TotMarkupValue / rs.RecordCount - 1, 2) & "%"
    grdTotal.TextMatrix(0, 4) = "Average Markup%: " & Round(TotMarkupValue / rs.RecordCount - 1, 2) & "%"
    grdTotal.TextMatrix(0, 5) = "Average GP%: " & Round(ToTGPValue / rs.RecordCount - 1, 2) & "%"
    grdTotal.TextMatrix(0, 6) = "Average GP%: " & Round(ToTGPValue / rs.RecordCount - 1, 2) & "%"
    rs.Close
    grdTotal.MergeRow(0) = True
    If grdMain.Rows > 1 Then
        grdMain.Row = 1
        grdMain.SetFocus
    End If
End Sub
Private Sub Sales_Analysis_Debtor()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    If cmb1.Text = "<All Debtors>" Then
        ActiveReadServer "Select Product_Code,Description,SUM(Sales_Tax) AS Sales_Tax, SUM(Line_Total) AS Line_Total " & _
        ",SUM(Qty) AS Qty," & _
        " SUM(Ave_Cost*Qty) AS Ave_Cost," & _
        " SUM(Line_Total / ((100 + Sales_Tax) / 100)) As Line_Total_Excl" & _
        " from Sales_Analysis_Debtor_View WHERE (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') " & _
        " Group by Product_Code,Description Order by Description"
    Else
        ActiveReadServer "Select Product_Code,Description,SUM(Sales_Tax) AS Sales_Tax, SUM(Line_Total) AS Line_Total " & _
        ",SUM(Qty) AS Qty," & _
        " SUM(Ave_Cost*Qty) AS Ave_Cost," & _
        " SUM(Line_Total / ((100 + Sales_Tax) / 100)) As Line_Total_Excl" & _
        " from Sales_Analysis_Debtor_View WHERE (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') " & _
        " and Account_No = '" & Trim(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 1)) & "' Group by Product_Code,Description Order by Description"
    End If
    
    grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = "0"
    grdTotal.TextMatrix(0, 3) = "0.00"
    grdTotal.TextMatrix(0, 4) = "0"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    ToTGPValue = 0
    TotTaxValue1 = 0
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Qty")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Ave_Cost"), "0.00")
        If Round(rs.Fields("Line_Total_Excl"), 3) = 0 Then
            grdMain.TextMatrix(grdMain.Rows - 1, 4) = "0%"
        Else
            grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(((rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")) / rs.Fields("Line_Total_Excl")) * 100, 2) & "%"
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Line_Total"), "0.00")
        TotTaxValue1 = TotTaxValue1 + (rs.Fields("Line_Total") - rs.Fields("Line_Total_Excl"))
        ToTGPValue = ToTGPValue + rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost"), "0.00")
        grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), "0.00")
        grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + Val(Replace(grdMain.TextMatrix(grdMain.Rows - 1, 4), "%", "")), 2)
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 6), "0.00")
        rs.MoveNext
    Wend
    If rs.RecordCount = 0 Then
        grdTotal.TextMatrix(0, 4) = "0%"
    Else
        grdTotal.TextMatrix(0, 4) = Round((ToTGPValue / (grdTotal.ValueMatrix(0, 6) - TotTaxValue1)) * 100, 3) & " %"
    End If
    rs.Close
    If grdTotal.ValueMatrix(0, 2) = 0 Then
        AveGP = 0
    Else
        AveGP = ToTGPValue / grdTotal.ValueMatrix(0, 2)
    End If
    grdTotal.TextMatrix(0, 1) = " Ave GP Value: " & Format(AveGP, "0.00")
    grdTotal.MergeRow(0) = True
    For i = 1 To grdMain.Rows - 1
        grdMain.RowHidden(i) = False
        If grdMain.ValueMatrix(i, 2) = 0 Then
            grdMain.TextMatrix(i, 7) = "Problem Child"
        Else
            grdMain.TextMatrix(i, 7) = Round(((grdMain.ValueMatrix(i, 5) / grdMain.ValueMatrix(i, 2)) / AveGP) * 100, 2)
            Select Case grdMain.ValueMatrix(i, 7)
                Case Is >= 100
                    grdMain.TextMatrix(i, 7) = "Star"
                Case Is > 66.66
                    grdMain.TextMatrix(i, 7) = "Cash Cow"
                Case Is > 33.33
                    grdMain.TextMatrix(i, 7) = "Dog"
                Case Is > 0
                    grdMain.TextMatrix(i, 7) = "Problem Child"
                Case 0
                    grdMain.TextMatrix(i, 7) = "Star"
                Case Is < 0
                    grdMain.TextMatrix(i, 7) = "Star"
            End Select
        End If
    Next i
    If grdMain.Rows > 1 Then
        grdMain.Row = 1
        grdMain.SetFocus
    End If
End Sub

Private Sub Sales_Analysis_Supplier()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    DepString = "%"
    If cmb1.Text = "<All Suppliers>" Then
        Suppstring = "%"
    Else
        If InStr(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2), "-") = 0 Then
            Suppstring = Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1)) & "%"
        Else
            Suppstring = Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1))
        End If
    End If
    
    ActiveReadServer "SELECT Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product') AS Description," & _
    "SUM(Sales_Journal.Sales_Tax) AS Sales_Tax, SUM(Sales_Journal.Line_Total) AS Line_Total, Sales_Journal.Department_No," & _
    "SUM(Sales_Journal.Qty) AS Qty, SUM(Sales_Journal.Ave_Cost*Sales_Journal.Qty) AS Ave_Cost," & _
    "SUM(Sales_Journal.Line_Total / ((100 + Sales_Journal.Sales_Tax) / 100)) As Line_Total_Excl " & _
    "FROM Sales_Journal LEFT OUTER JOIN " & _
    "Products ON Sales_Journal.Product_Code = Products.Product_Code " & _
    "WHERE (Sales_Journal.Function_Key IN (7))  and Sales_Journal.Department_No like '" & DepString & "' AND ((ISNULL(Sales_Journal.Extra, '') = '' or Sales_Journal.Extra='Return Item')) and " & _
    "(Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " GROUP BY Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product'), Sales_Journal.Department_No " & _
    "HAVING (Sales_Journal.Product_Code in (Select Product_Code from Supplier_Links where Supplier_No like '" & Suppstring & "'))"
    grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = "0"
    grdTotal.TextMatrix(0, 3) = "0.00"
    grdTotal.TextMatrix(0, 4) = "0"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    ToTGPValue = 0
    TotTaxValue1 = 0
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Qty")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Ave_Cost"), "0.00")
        If Round(rs.Fields("Line_Total_Excl"), 3) = 0 Then
            grdMain.TextMatrix(grdMain.Rows - 1, 4) = "0%"
        Else
            grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(((rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")) / rs.Fields("Line_Total_Excl")) * 100, 2) & "%"
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Line_Total"), "0.00")
        TotTaxValue1 = TotTaxValue1 + (rs.Fields("Line_Total") - rs.Fields("Line_Total_Excl"))
        ToTGPValue = ToTGPValue + rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost"), "0.00")
        grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), "0.00")
        grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + Val(Replace(grdMain.TextMatrix(grdMain.Rows - 1, 4), "%", "")), 2)
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 6), "0.00")
        rs.MoveNext
    Wend
    If rs.RecordCount = 0 Then
        grdTotal.TextMatrix(0, 4) = "0%"
    Else
        grdTotal.TextMatrix(0, 4) = Round((ToTGPValue / (grdTotal.ValueMatrix(0, 6) - TotTaxValue1)) * 100, 3) & " %"
    End If
    rs.Close
    If grdTotal.ValueMatrix(0, 2) = 0 Then
        AveGP = 0
    Else
        AveGP = ToTGPValue / grdTotal.ValueMatrix(0, 2)
    End If
    grdTotal.TextMatrix(0, 1) = " Ave GP Value: " & Format(AveGP, "0.00")
    grdTotal.MergeRow(0) = True
    For i = 1 To grdMain.Rows - 1
        grdMain.RowHidden(i) = False
        If grdMain.ValueMatrix(i, 2) = 0 Then
            grdMain.TextMatrix(i, 7) = "Problem Child"
        Else
            grdMain.TextMatrix(i, 7) = Round(((grdMain.ValueMatrix(i, 5) / grdMain.ValueMatrix(i, 2)) / AveGP) * 100, 2)
            Select Case grdMain.ValueMatrix(i, 7)
                Case Is >= 100
                    grdMain.TextMatrix(i, 7) = "Star"
                Case Is > 66.66
                    grdMain.TextMatrix(i, 7) = "Cash Cow"
                Case Is > 33.33
                    grdMain.TextMatrix(i, 7) = "Dog"
                Case Is > 0
                    grdMain.TextMatrix(i, 7) = "Problem Child"
                Case 0
                    grdMain.TextMatrix(i, 7) = "Star"
                Case Is < 0
                    grdMain.TextMatrix(i, 7) = "Star"
            End Select
        End If
    Next i
    If grdMain.Rows > 1 Then
        grdMain.Row = 1
        grdMain.SetFocus
    End If
End Sub
Private Sub Room_Sales()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    Select Case cmb3.Text
        Case "No Sales": FunctionKey = 6
        Case "Cash Sales": FunctionKey = 9
        Case "Card Sales": FunctionKey = 10
        Case "Voucher Sales": FunctionKey = 11
        Case "Charge Sales": FunctionKey = 12
        Case "Loyalty Sales": FunctionKey = 13
    End Select
    
    If cmb1.Text <> "<All Users>" Then
        If cmb3.Text = "<All Transactions>" Then
            ActiveReadServer "Select * from Sales_Journal_View where isnull(Room_No,0) <> 0 and User_No= " & Trim(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
        Else
            ActiveReadServer "Select * from Sales_Journal_View where isnull(Room_No,0) <> 0 and User_No= " & Trim(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') and Function_Key = " & FunctionKey & " order by Date_Time"
        End If
    Else
        If cmb3.Text = "<All Transactions>" Or cmb3.Text = "" Then
            ActiveReadServer "Select * from Sales_Journal_View where isnull(Room_No,0) <> 0 and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
        Else
            Select Case cmb3.Text
                Case "Room Accomodation"
                    ActiveReadServer "Select * from Sales_Journal_View where isnull(Room_No,0) <> 0 and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') " & _
                    " and Invoice_No in (Select invoice_No from Sales_Journal where Function_key=20) " & _
                    " order by Date_Time"
                Case "Other"
                    ActiveReadServer "Select * from Sales_Journal_View where isnull(Room_No,0) <> 0 and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') " & _
                    " and Invoice_No in (Select invoice_No from Sales_Journal where Function_key<>20 and Function_Key =7) " & _
                    " order by Date_Time"
            End Select
            
        End If
    End If
    grdTotal.TextMatrix(0, 0) = " Totals:"
    grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 5) = "0"
    grdTotal.TextMatrix(0, 6) = "0"
    grdTotal.TextMatrix(0, 7) = "0"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY HH:mm:ss AM/PM")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("User_No") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Workstation_No") & " - " & rs.Fields("Workstation_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Invoice_No"), "00000")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Key_Desc") & " - Room: " & Trim(rs.Fields("Room_No"))
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = Format(rs.Fields("Line_Total"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Sales_Tax"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total") - rs.Fields("Sales_Tax"), "0.00")
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(grdMain.Rows - 1, 7), "0.00")
        If rs.Fields("Function_Key") = 6 Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 7) = &HD7FDD9
        End If
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub
Private Sub Stock_Movement_Quantities()
    On Error GoTo trap
    Screen.MousePointer = 13
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    grdTotal.TextMatrix(0, 2) = "0"
    grdTotal.TextMatrix(0, 3) = "0"
    grdTotal.TextMatrix(0, 4) = "0"
    grdTotal.TextMatrix(0, 5) = "0"
    grdTotal.TextMatrix(0, 6) = "0"
    grdTotal.TextMatrix(0, 7) = "0"
    grdTotal.TextMatrix(0, 8) = "0"
    grdTotal.TextMatrix(0, 9) = "0"
    grdTotal.TextMatrix(0, 10) = "0"
    grdTotal.TextMatrix(0, 11) = "0"
    grdTotal.TextMatrix(0, 12) = "0"
    grdMain.Rows = 1
    grdMain.SetFocus
    DoEvents
    ActiveReadServer3 "SELECT Products.Product_Code,Description From Products LEFT OUTER JOIN Quantities ON Products.Product_Code = Quantities.Product_Code where Location_No like '" & LocString & "' and Department_No like '" & DeptString & "' and Stock_Item=1 GROUP BY Products.Product_Code,Description order by Description"
    While Not rs3.EOF
        If rs3.Fields("Product_code") & "" <> "" Then
            grdMain.Rows = grdMain.Rows + 1
            For i = 2 To grdMain.Cols - 1
               grdMain.TextMatrix(grdMain.Rows - 1, i) = "0"
            Next i
            ActiveReadServer1 "Exec Stock_Move '" & DeptString & "', '" & LocString & "', '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "', '" & mthViewEnd.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "', '" & rs3.Fields("Product_code") & "'"
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs1.Fields("Product_code")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs1.Fields("Description")
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = Round(Val(rs1.Fields("Qty_Received") & ""), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Round(Val(rs1.Fields("Qty_Trans_Out") & ""), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 7) = Round(Val(rs1.Fields("Qty_Trans_In") & ""), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 9) = Round(Val(rs1.Fields("Qty_Consumed") & ""), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 11) = Round(Val(rs1.Fields("Variance") & ""), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 12) = Round(Val(rs1.Fields("Stock_on_Hand") & ""), 3)
            rs1.Close
            grdMain.TextMatrix(grdMain.Rows - 1, 10) = Round(grdMain.ValueMatrix(grdMain.Rows - 1, 12) - grdMain.ValueMatrix(grdMain.Rows - 1, 11), 3)
            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Round(grdMain.ValueMatrix(grdMain.Rows - 1, 10) + grdMain.ValueMatrix(grdMain.Rows - 1, 9) + grdMain.ValueMatrix(grdMain.Rows - 1, 6) - grdMain.ValueMatrix(grdMain.Rows - 1, 7) - grdMain.ValueMatrix(grdMain.Rows - 1, 3), 3)
            grdTotal.TextMatrix(0, 2) = Round(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), 3)
            grdTotal.TextMatrix(0, 3) = Round(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), 3)
            grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), 3)
            grdTotal.TextMatrix(0, 5) = Round(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), 3)
            grdTotal.TextMatrix(0, 6) = Round(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), 3)
            grdTotal.TextMatrix(0, 7) = Round(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(grdMain.Rows - 1, 7), 3)
            grdTotal.TextMatrix(0, 8) = Round(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), 3)
            grdTotal.TextMatrix(0, 9) = Round(grdTotal.ValueMatrix(0, 9) + grdMain.ValueMatrix(grdMain.Rows - 1, 9), 3)
            grdTotal.TextMatrix(0, 10) = Round(grdTotal.ValueMatrix(0, 10) + grdMain.ValueMatrix(grdMain.Rows - 1, 10), 3)
            grdTotal.TextMatrix(0, 11) = Round(grdTotal.ValueMatrix(0, 11) + grdMain.ValueMatrix(grdMain.Rows - 1, 11), 3)
            grdTotal.TextMatrix(0, 12) = Round(grdTotal.ValueMatrix(0, 12) + grdMain.ValueMatrix(grdMain.Rows - 1, 12), 3)
            grdMain.ShowCell grdMain.Rows - 1, 0
            DoEvents
        End If
        rs3.MoveNext
    Wend
    grdTotal.TextMatrix(0, 1) = " Products = " & grdMain.Rows - 1
    rs3.Close
    If grdMain.Rows > 1 Then
         grdMain.Row = 1
         grdMain.ShowCell 1, 0
         grdMain.SetFocus
    End If
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
trap:
    rs3.Close
    rs1.Close
    Screen.MousePointer = 0
    On Error GoTo 0
End Sub

Private Sub Boktime_Click()
'             Pictime.Left = 9835
'             Pictime.top = 660
'             Lbltimeback.Left = 9860
'             Lbltimeback.top = 660
'             Lbltime.Left = 9900
'             Lbltime.top = 730
            Lbltime.Caption = Format(TStart.Value, "Medium Time") & " to " & Format(TEnd.Value, "Medium Time")
 
'Pictime.Visible = False
'Image8.Visible = False
'Boktime.Visible = False
'TStart.Visible = False
'TEnd.Visible = False

Pictimeback.Visible = False


End Sub

Private Sub Buttime_Click()


Pictimeback.Visible = True



End Sub

Private Sub ButtonEx1_Click()
    Select Case ButtonEx1.Value
        Case 0
            picDate.Visible = True
        Case 1
            picDate.Visible = False
            If picDate.Visible = False Then Selection_Change
    End Select
End Sub



Private Sub cmb1_Change()
    Selection_Change
End Sub
Private Sub Room_Accounts()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    grdMain.Rows = 1
    Select Case cmb1.Text
        Case "<All Reservations>"
            ActiveReadServer "Select * from Res_View " & _
            " where Res_Type <> 0 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Confirmed Bookings"
            ActiveReadServer "Select * from Res_View" & _
            " where Res_Type = 1 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Guests Checked In"
            ActiveReadServer "Select * from Res_View" & _
            " where Res_Type = 2 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
        Case "Guests Checked Out"
            ActiveReadServer "Select * from Res_View" & _
            " where Res_Type = 3 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
            " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
            " group by Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type order by Arrive_Date"
    End Select
    grdTotal.TextMatrix(0, 0) = " Totals:"
    grdTotal.TextMatrix(0, 1) = " Room Accounts = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Room Accounts = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Room Accounts = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = " Room Accounts = " & rs.RecordCount
    grdTotal.TextMatrix(0, 5) = " Room Accounts = " & rs.RecordCount
    grdTotal.TextMatrix(0, 6) = "0.00"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Room_No") & " - " & rs.Fields("Description")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Guest_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Arrive_Date"), "dd MMM yyyy ddd")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Depart_Date"), "dd MMM yyyy ddd")
        Select Case rs.Fields("Res_Type")
            Case 1: grdMain.TextMatrix(grdMain.Rows - 1, 4) = "Confirmed"
            Case 2: grdMain.TextMatrix(grdMain.Rows - 1, 4) = "Checked In"
            Case 3: grdMain.TextMatrix(grdMain.Rows - 1, 4) = "Checked Out"
        End Select
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Val(rs.Fields("Nights"))
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(Val(rs.Fields("Balance") & ""), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = rs.Fields("Res_No")
        Select Case rs.Fields("Res_Type")
            Case 0: grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 6) = &HC0FFFF
            Case 1: grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 6) = &HC0FFC0
            Case 2: grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 6) = &HFDE0DF
            Case 3: grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 6) = &HC0C0FF
        End Select
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
    If grdMain.Rows > 1 Then grdMain.Row = 1
End Sub
Private Sub Deposits_Paid()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    grdMain.Rows = 1
    Select Case cmb3.Text
        Case "Cash"
            Extra = " Tender_Type='Cash' and"
        Case "Voucher"
            Extra = " Tender_Type='Voucher' and"
        Case "Charge"
            Extra = " Tender_Type='Charge' and"
        Case "Card"
            Extra = " Tender_Type='Card' and"
        Case "EFT"
            Extra = " Tender_Type='EFT' and"
        Case Else
            Extra = ""
    End Select
    If cmb2.Text = "Payments Received" Then
        Select Case cmb1.Text
            Case "<All Reservations>"
                ActiveReadServer "Select Date_Time,Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type,Credit as Receipts,Tender_Type from Res_View1 " & _
                " where Transaction_Type ='Receipt' and" & Extra & " ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
                " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
                " order by Arrive_Date"
            Case "Guests Checked In"
                ActiveReadServer "Select Date_Time,Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type,Credit as Receipts,Tender_Type from Res_View1 " & _
                " where Transaction_Type ='Receipt' and" & Extra & " Res_Type = 2 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
                " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
                " order by Arrive_Date"
            Case "Guests Checked Out"
                ActiveReadServer "Select Date_Time,Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type,Credit as Receipts,Tender_Type from Res_View1 " & _
                " where Transaction_Type ='Receipt' and" & Extra & " Res_Type = 3 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
                " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
                " order by Arrive_Date"
        End Select
    Else
        Select Case cmb1.Text
            Case "<All Reservations>"
                ActiveReadServer "Select Date_Time,Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type,Credit as Deposit,Tender_Type from Res_View1 " & _
                " where Transaction_Type ='Deposit' and" & Extra & " ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
                " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
                " order by Arrive_Date"
            Case "Confirmed Bookings"
                ActiveReadServer "Select Date_Time,Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type,Credit as Deposit,Tender_Type from Res_View1 " & _
                " where Transaction_Type ='Deposit' and" & Extra & " Res_Type = 1 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
                " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
                " order by Arrive_Date"
            Case "Guests Checked In"
                ActiveReadServer "Select Date_Time,Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type,Credit as Deposit,Tender_Type from Res_View1 " & _
                " where Transaction_Type ='Deposit' and" & Extra & " Res_Type = 2 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
                " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
                " order by Arrive_Date"
            Case "Guests Checked Out"
                ActiveReadServer "Select Date_Time,Res_No,Room_No,Description,Arrive_Date,Depart_Date,Guest_Name,Tel_No,Nights,Balance,Res_Type,Credit as Deposit,Tender_Type from Res_View1 " & _
                " where Transaction_Type ='Deposit' and" & Extra & " Res_Type = 3 and ((Arrive_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Arrive_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "')" & _
                " or (Depart_Date > '" & DateAdd("D", -1, mthViewStart.Value) & "'and Depart_Date < '" & DateAdd("D", 1, mthViewEnd.Value) & "'))" & _
                " order by Arrive_Date"
        End Select
    End If
    grdTotal.TextMatrix(0, 0) = " Totals:"
    grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 5) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy ddd")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Room_No") & " - " & rs.Fields("Description")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Guest_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Arrive_Date"), "dd MMM yyyy ddd")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Depart_Date"), "dd MMM yyyy ddd")
        Select Case rs.Fields("Res_Type")
            Case 1: grdMain.TextMatrix(grdMain.Rows - 1, 5) = "Confirmed"
            Case 2: grdMain.TextMatrix(grdMain.Rows - 1, 5) = "Checked In"
            Case 3: grdMain.TextMatrix(grdMain.Rows - 1, 5) = "Checked Out"
        End Select
        If cmb2.Text = "Payments Received" Then
            grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(Val(rs.Fields("Receipts") & ""), "0.00")
        Else
            grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(Val(rs.Fields("Deposit") & ""), "0.00")
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = rs.Fields("Tender_Type") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = rs.Fields("Res_No")
        Select Case rs.Fields("Res_Type")
            Case 0: grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 7) = &HC0FFFF
            Case 1: grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 7) = &HC0FFC0
            Case 2: grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 7) = &HFDE0DF
            Case 3: grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 7) = &HC0C0FF
        End Select
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
    If grdMain.Rows > 1 Then grdMain.Row = 1
End Sub
Private Sub Sales_Journal()
    Screen.MousePointer = 11
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    Select Case cmb3.Text
        Case "No Sales": FunctionKey = 6
        Case "Cash Sales": FunctionKey = 9
        Case "Card Sales": FunctionKey = 10
        Case "Voucher Sales": FunctionKey = 11
        Case "Charge Sales": FunctionKey = 12
        Case "Loyalty Sales": FunctionKey = 13
    End Select
    
    If cmb1.Text <> "<All Users>" Then
        If cmb3.Text = "<All Transactions>" Then
            ActiveReadServer "Select * from Sales_Journal_View where User_No= " & Trim(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & " and Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "' order by Date_Time"
        Else
            ActiveReadServer "Select * from Sales_Journal_View where User_No= " & Trim(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & " and Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "' and Function_Key = " & FunctionKey & " order by Date_Time"
        End If
    Else
        If cmb3.Text = "<All Transactions>" Or cmb3.Text = "" Then
            ActiveReadServer "Select * from Sales_Journal_View where Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "' order by Date_Time"
        Else
            ActiveReadServer "Select * from Sales_Journal_View where Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "' and Function_Key = " & FunctionKey & " order by Date_Time"
        End If
    End If
    grdTotal.TextMatrix(0, 0) = " Totals:"
    grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 5) = "0"
    grdTotal.TextMatrix(0, 6) = "0"
    grdTotal.TextMatrix(0, 7) = "0"
    grdTotal.TextMatrix(0, 8) = "0"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY HH:mm:ss AM/PM")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("User_No") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
        If rs.Fields("Workstation_Name") = "Table No: 0.0" Then
            grdMain.TextMatrix(grdMain.Rows - 1, 2) = "Workstation: " & rs.Fields("Workstation_No")
        Else
            If rs.Fields("Key_Desc") = "No Sale" Then
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Workstation_No") & " - No Sale"
            Else
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Workstation_No") & " - " & rs.Fields("Workstation_Name")
            End If
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = String(8 - Len(rs.Fields("Invoice_No")), "0") & rs.Fields("Invoice_No")
        If rs.Fields("Key_Desc") = "Charge Sale" Then
            grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Key_Desc") & " - " & rs.Fields("Debtor_Name")
        Else
            grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Key_Desc")
        End If
        Dim rsval As String
        rsval = Format(rs.Fields("Covers"), "0")
        If rsval = "" Then rsval = 0
        
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = Format(rs.Fields("Line_Total"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Sales_Tax"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total") - rs.Fields("Sales_Tax"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = rsval
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(grdMain.Rows - 1, 7), "0.00")
        grdTotal.TextMatrix(0, 8) = Format(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), "0")
        If rs.Fields("Function_Key") = 6 Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 7) = &HD7FDD9
        End If
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
    Screen.MousePointer = 0
End Sub
Private Sub User_Shift_List()
    On Error GoTo trap
    Dim StartTime As Date
    Dim StopTime As Date
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    
    Debug.Print "Select Function_key,User_No,Date_Time," & _
        " isnull((Select First_Name+' ' +Last_Name from Users where Users.User_No=User_Journal.User_No),'Deleted User') as Users from User_Journal" & _
        " where function_key in (3,4) " & _
        " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
        " order by User_No,Date_Time,Function_Key"
    
    
    
    If cmb1.Text <> "<All Users>" Then
        ActiveReadServer "Select Function_key,User_No,Date_Time," & _
        " (Select First_Name+' ' +Last_Name from Users where Users.User_No=User_Journal.User_No) as Users from User_Journal" & _
        " where function_key in (3,4) and User_No = " & Trim(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & _
        " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
        " order by User_No,Date_Time,Function_Key"
    Else
        ActiveReadServer "Select Function_key,User_No,Date_Time," & _
        " isnull((Select First_Name+' ' +Last_Name from Users where Users.User_No=User_Journal.User_No),'Deleted User') as Users from User_Journal" & _
        " where function_key in (3,4) " & _
        " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
        " order by User_No,Date_Time,Function_Key"
    End If
    grdTotal.TextMatrix(0, 0) = ""
    grdTotal.TextMatrix(0, 1) = "0.00"
    grdTotal.TextMatrix(0, 2) = "0.00"
    grdTotal.TextMatrix(0, 3) = "0.00"
    grdTotal.TextMatrix(0, 4) = "0.00"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdMain.Rows = 1
    If rs.RecordCount > 0 And rs.Fields("Function_Key") = 4 Then
        rs.MoveNext
    End If
    newline = False
    Total = 0
    While Not rs.EOF
        If rs.Fields("Function_Key") = 3 Then
            StartTime = rs.Fields("Date_Time")
        Else
            StopTime = rs.Fields("Date_Time")
            newline = True
        End If
        If newline = True Then
            grdMain.Rows = grdMain.Rows + 1
            grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY (DDD)")
            grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Users") & ""
            '????????????
            If grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Ria" Then Stop
            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(StartTime, "DD MMM YYYY   HH:MM:SS")
            grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(StopTime, "DD MMM YYYY   HH:MM:SS")
            grdMain.TextMatrix(grdMain.Rows - 1, 4) = "0.00"
            grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(Int(DateDiff("n", StartTime, StopTime) / 60), "00") & " Hrs " & DateDiff("n", StartTime, StopTime) - (Int(DateDiff("n", StartTime, StopTime) / 60) * 60) & " Min"
            Total = Total + DateDiff("n", StartTime, StopTime)
            newline = False
        End If
        rs.MoveNext
    Wend
    grdTotal.TextMatrix(0, 0) = ""
    grdTotal.TextMatrix(0, 1) = "Total Shifts: " & grdMain.Rows - 1
    grdTotal.TextMatrix(0, 2) = "Total Shifts: " & grdMain.Rows - 1
    grdTotal.TextMatrix(0, 3) = ""
    grdTotal.TextMatrix(0, 4) = "0.00"
    grdTotal.TextMatrix(0, 5) = Format(Int(Total / 60), "00") & " Hrs " & Total - (Int(Total / 60) * 60) & " Min"
    grdTotal.MergeRow(0) = True
    rs.Close
top:
    Total = 0
    For i = 1 To grdMain.Rows - 1
        If InStr(grdMain.TextMatrix(i, 5), "-") <> 0 Then
            grdMain.RemoveItem i
            GoTo top
        End If
        Total = Total + DateDiff("n", grdMain.TextMatrix(i, 2), grdMain.TextMatrix(i, 3))
    Next i
    grdTotal.TextMatrix(0, 5) = Format(Int(Total / 60), "00") & " Hrs " & Total - (Int(Total / 60) * 60) & " Min"
    grdMain.ColHidden(4) = True
    grdTotal.ColHidden(4) = True
    On Error GoTo 0
    Exit Sub
trap:
    Resume Next
End Sub
Private Sub Receive_on_Account()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    grdMain.Rows = 1
    Select Case cmb1.Text
        Case "Cash"
            Extra = " Tender_Type='Cash' and"
        Case "Voucher"
            Extra = " Tender_Type='Voucher' and"
        Case "Charge"
            Extra = " Tender_Type='Charge' and"
        Case "Card"
            Extra = " Tender_Type='Card' and"
        Case "EFT"
            Extra = " Tender_Type='EFT' and"
        Case Else
            Extra = ""
    End Select
    ActiveReadServer "Select * from Debt_View " & _
    " where Transaction_Type ='Receipt' and" & Extra & " " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
    grdTotal.TextMatrix(0, 0) = " Totals:"
    grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = "0.00"
    grdTotal.TextMatrix(0, 5) = "0.00"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "dd MMM yyyy ddd")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Account_No")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Debtor_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("Business_Tel")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Credit"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = rs.Fields("Tender_Type") & ""
        grdTotal.TextMatrix(0, 4) = Format(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), "0.00")
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
    If grdMain.Rows > 1 Then grdMain.Row = 1
End Sub
Private Sub Sales_Correct()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    Select Case cmb1.Text
        Case "Item Corrects": Extra = "Corr"
        Case "Voids": Extra = "Void"
        Case "Returns": Extra = "Return Item"
        Case Else: Extra = "%"
    End Select
    If cmb2.Text = "Discounted Sales" Then
        If cmb1.Text = "<All Users>" Then
            ActiveReadServer "Select * from Discount_View where " & _
            " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
        Else
            ActiveReadServer "Select * from Discount_View where User_No = " & Val(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & " and " & _
            " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
        End If
    Else
        If cmb3.Text = "<All Users>" Then
            ActiveReadServer "Select * from Corr_View where Extra like '" & Extra & "'" & _
            " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
        Else
            ActiveReadServer "Select * from Corr_View where Extra like '" & Extra & "' and User_No = " & Val(Mid(cmb3.Text, 1, InStr(cmb3.Text, "-") - 1)) & _
            " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
        End If
    End If
    grdTotal.TextMatrix(0, 0) = " Totals:"
    grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = "0"
    grdTotal.TextMatrix(0, 5) = "0"
    grdTotal.TextMatrix(0, 6) = ""
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY HH:mm:ss AM/PM")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Description")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = String(8 - Len(rs.Fields("Invoice_No")), "0") & rs.Fields("Invoice_No")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Qty")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total"), "0.00") & " "
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = rs.Fields("User")
        grdTotal.TextMatrix(0, 4) = Format(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), "0.00")
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub
Private Sub Stock_Takes()
    On Error Resume Next
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    '********************************************************
'    If cbm1.Text = "<Extra Stock Functions>" Then Exit Sub
'    If cbm1.Text = "Stock not counted during stocktake" Then
'    Stock_not_counted
'    End If
'    Exit Sub
   '*********************************************************
    
    
    If cmb1.Text = "<All Locations>" Then
        ActiveReadServer "Select * from Stock_Take_View where " & _
        " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
    Else
        ActiveReadServer "Select * from Stock_Take_View where Location_No=" & Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1) & _
        " and ((Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')) order by Date_Time"
    End If
    grdTotal.TextMatrix(0, 0) = " Totals:"
    
    RowCount = rs.RecordCount
    grdTotal.TextMatrix(0, 4) = "0"
    grdTotal.TextMatrix(0, 5) = "0"
    grdTotal.TextMatrix(0, 6) = "0"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY HH:mm:ss AM/PM")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Take_No")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Location")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = "Stock Take"
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Stock_on_Hand"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Variance"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Counted"), "0.00")
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    
    ActiveReadServer "Select SUM(Line_Total) AS Line_Total" & _
            " From Sales_Journal" & _
            " Where (Function_Key in(7,20))" & _
            " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
            " and Location like '" & LocString & "%'" & _
            " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    Turnover = rs.Fields("Line_Total")
    rs.Close
    Stock_Value = 0
    ActiveReadServer "Select * from Stock_Value" & _
    " WHERE Location_No like '" & LocString & "'"
    While Not rs.EOF
        Stock_Value = Stock_Value + rs.Fields("Stock_Value")
        rs.MoveNext
    Wend
    rs.Close
    
    Variance = grdTotal.TextMatrix(0, 5)
    grdTotal.TextMatrix(0, 1) = " Transactions = " & RowCount & " (Percentage of Stock Holding: " & Round((Variance / Stock_Value) * 100, 3) & "%)"
    grdTotal.TextMatrix(0, 2) = " Transactions = " & RowCount & " (Percentage of Stock Holding: " & Round((Variance / Stock_Value) * 100, 3) & "%)"
    grdTotal.TextMatrix(0, 3) = " Transactions = " & RowCount & " (Percentage of Stock Holding: " & Round((Variance / Stock_Value) * 100, 3) & "%)"
    grdTotal.TextMatrix(0, 5) = Variance & " (Percentage of Revenue: " & Round((Variance / Turnover) * 100, 3) & "%)"
    grdTotal.TextMatrix(0, 4) = Variance & " (Percentage of Revenue: " & Round((Variance / Turnover) * 100, 3) & "%)"
    grdTotal.TextMatrix(0, 6) = Variance & " (Percentage of Revenue: " & Round((Variance / Turnover) * 100, 3) & "%)"
    
    grdTotal.MergeRow(0) = True
    On Error GoTo 0
End Sub
'Private Sub Stock_not_counted()
'Exit Sub
'        ActiveReadServer "Select * from Stock_Take_View where " & _
'        " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
'    Else
'        ActiveReadServer "Select * from Stock_Take_View where Location_No=" & Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1) & _
'        " and ((Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')) order by Date_Time"
'    End If
'    grdTotal.TextMatrix(0, 0) = " Totals:"
'
'    RowCount = rs.RecordCount
'    grdTotal.TextMatrix(0, 4) = "0"
'    grdTotal.TextMatrix(0, 5) = "0"
'    grdTotal.TextMatrix(0, 6) = "0"
'    While Not rs.EOF
'        grdMain.Rows = grdMain.Rows + 1
'        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY HH:mm:ss AM/PM")
'        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Take_No")
'        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Location")
'        grdMain.TextMatrix(grdMain.Rows - 1, 3) = "Stock Take"
'        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Stock_on_Hand"), "0.00")
'        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Variance"), "0.00")
'        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Counted"), "0.00")
'        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
'        rs.MoveNext
'    Wend
'    rs.Close
'    If cmb1.Text <> "<All Locations>" Then
'        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
'    Else
'        LocString = "%"
'    End If
'
'    ActiveReadServer "Select SUM(Line_Total) AS Line_Total" & _
'            " From Sales_Journal" & _
'            " Where (Function_Key in(7,20))" & _
'            " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
'            " and Location like '" & LocString & "%'" & _
'            " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
'    Turnover = rs.Fields("Line_Total")
'    rs.Close
'    Stock_Value = 0
'    ActiveReadServer "Select * from Stock_Value" & _
'    " WHERE Location_No like '" & LocString & "'"
'    While Not rs.EOF
'        Stock_Value = Stock_Value + rs.Fields("Stock_Value")
'        rs.MoveNext
'    Wend
'    rs.Close
'
'    Variance = grdTotal.TextMatrix(0, 5)
'    grdTotal.TextMatrix(0, 1) = " Transactions = " & RowCount & " (Percentage of Stock Holding: " & Round((Variance / Stock_Value) * 100, 3) & "%)"
'    grdTotal.TextMatrix(0, 2) = " Transactions = " & RowCount & " (Percentage of Stock Holding: " & Round((Variance / Stock_Value) * 100, 3) & "%)"
'    grdTotal.TextMatrix(0, 3) = " Transactions = " & RowCount & " (Percentage of Stock Holding: " & Round((Variance / Stock_Value) * 100, 3) & "%)"
'    grdTotal.TextMatrix(0, 5) = Variance & " (Percentage of Revenue: " & Round((Variance / Turnover) * 100, 3) & "%)"
'    grdTotal.TextMatrix(0, 4) = Variance & " (Percentage of Revenue: " & Round((Variance / Turnover) * 100, 3) & "%)"
'    grdTotal.TextMatrix(0, 6) = Variance & " (Percentage of Revenue: " & Round((Variance / Turnover) * 100, 3) & "%)"
'
'    grdTotal.MergeRow(0) = True
'    On Error GoTo 0
'End Sub




Private Sub Stock_Summary()
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    grdTotal.TextMatrix(0, 2) = "0"
    grdTotal.TextMatrix(0, 3) = "0.00"
    grdTotal.TextMatrix(0, 4) = "0.00"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    ActiveReadServer "Select * from Stock_Summary_View where " & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'"
    grdTotal.TextMatrix(0, 0) = " Departments = " & rs.RecordCount
    grdTotal.TextMatrix(0, 1) = " Departments = " & rs.RecordCount
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Department_No") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Dept_Name") & " - " & rs.Fields("Loc_Name") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Round(Val(rs.Fields("Stock_on_Hand") & ""), 3)
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Ave_Cost") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Stock_Value") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Sales_Tax") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Total_Incl") & "", "0.00")
        grdTotal.TextMatrix(0, 2) = Round(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), 3)
        grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
        grdTotal.TextMatrix(0, 4) = Format(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), "0.00")
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub
Private Sub Payout_Journal()
    
    '"(Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" &
    
    
    
    
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    If cmb1.Text <> "<All Suppliers>" Then
        If cmb3.Text = "<All Users>" Then
            ActiveReadServer "Select * from  Payout_View Where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and Account_No ='" & Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1)) & "' order by Date_Time"
        Else
            ActiveReadServer "Select * from  Payout_View Where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and Account_No ='" & Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1)) & "' and User_No = " & Val(Mid(cmb3.Text, 1, InStr(cmb3.Text, "-") - 1)) & " order by Date_Time"
        End If
    Else
        If cmb3.Text = "<All Users>" Or cmb3.Text = "" Then
            ActiveReadServer "Select * from  Payout_View Where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' order by Date_Time"
        Else
            ActiveReadServer "Select * from  Payout_View Where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and User_No = " & Val(Mid(cmb3.Text, 1, InStr(cmb3.Text, "-") - 1)) & " order by Date_Time"
        End If
    End If
    grdTotal.TextMatrix(0, 0) = " Totals:"
    grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = "0.00"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time") & "", "YYYY-MM-DD hh:mm")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Supplier") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("User") & ""
        If rs.Fields("Ref_No") & "" <> "" Then
            grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Payment_No"), "000000") & " - " & "Payout" & " - " & rs.Fields("Ref_No")
        Else
            grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Payment_No"), "000000") & " - " & "Payout"
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Debit") & "", "0.00")
        grdTotal.TextMatrix(0, 4) = Format(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub
Private Sub User_Commission_List() '  commision
    On Error Resume Next
    Dim StartTime As Date
    Dim StopTime As Date
    ActiveUpdateServer "Delete from Comm_Temp"
    DoEvents
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    
    
   
    If cmb1.Text = "<All Users>" Then
        ActiveReadServer "SELECT convert(nvarchar(10),User_No) + ' - '+isnull((Select User_Name from Users where Sales_Journal.User_No = Users.User_No),'Deleted User') as User_Name,isnull((Select Comm1 from Users where Sales_Journal.User_No = Users.User_No),0) as Comm," & _
        " Case Function_Key WHEN 9 THEN 'Cash' WHEN 10 THEN 'Card' WHEN 11 THEN 'Voucher' WHEN 12 THEN 'Charge' WHEN 13 THEN 'Loyalty' END As Sale_Type,SUM(Sales_Tax)AS Sales_Tax,SUM(Line_Total) AS Line_Total From Sales_Journal" & _
        " WHERE (Function_Key IN ( 9, 10, 11, 12, 13))  and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') GROUP BY Function_Key,User_No" & _
        " order by User_No"
    Else
        ActiveReadServer "SELECT convert(nvarchar(10),User_No) + ' - '+isnull((Select User_Name from Users where Sales_Journal.User_No = Users.User_No),'Deleted User') as User_Name,isnull((Select Comm1 from Users where Sales_Journal.User_No = Users.User_No),0) as Comm," & _
        " Case Function_Key WHEN 9 THEN 'Cash' WHEN 10 THEN 'Card' WHEN 11 THEN 'Voucher' WHEN 12 THEN 'Charge' WHEN 13 THEN 'Loyalty' END As Sale_Type,SUM(Sales_Tax)AS Sales_Tax,SUM(Line_Total) AS Line_Total From Sales_Journal" & _
        " WHERE (Function_Key IN ( 9, 10, 11, 12, 13))  and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') " & _
        " and User_No = '" & Val(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & "' GROUP BY Function_Key,User_No" & _
        " order by User_No"
    End If
    grdTotal.TextMatrix(0, 1) = "0.00"
    grdTotal.TextMatrix(0, 2) = "0.00"
    grdTotal.TextMatrix(0, 3) = "0.00"
    grdTotal.TextMatrix(0, 4) = "0.00"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    grdTotal.TextMatrix(0, 8) = "0.00"
    grdTotal.TextMatrix(0, 9) = "0.00"
    grdMain.Rows = 1
    If rs.RecordCount > 0 Then
        grdMain.Rows = grdMain.Rows + 1
    End If
    newline = True
    While Not rs.EOF
        User_Name = rs.Fields("User_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("User_Name")
        Select Case rs.Fields("Sale_Type")
            Case "Cash"
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = Format(rs.Fields("Line_Total"), "0.00")
            Case "Card"
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Line_Total"), "0.00")
            Case "Voucher"
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Line_Total"), "0.00")
            Case "Charge"
                grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Line_Total"), "0.00")
        End Select
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 1), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 2), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 4), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 1) + grdMain.ValueMatrix(grdMain.Rows - 1, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 6) + Format(rs.Fields("Sales_Tax"), "0.00"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = grdMain.ValueMatrix(grdMain.Rows - 1, 5) - grdMain.ValueMatrix(grdMain.Rows - 1, 6)
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = Round(rs.Fields("Comm"), 2) & "%"
        grdMain.TextMatrix(grdMain.Rows - 1, 9) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 7) - (grdMain.ValueMatrix(grdMain.Rows - 1, 7) * ((100 - rs.Fields("Comm")) / 100)), "0.00")
        rs.MoveNext
        If Not rs.EOF Then
            If User_Name <> rs.Fields("User_Name") Then
                grdMain.Rows = grdMain.Rows + 1
            End If
        End If
    Wend
    rs.Close
    For i = 1 To grdMain.Rows - 1
        grdTotal.TextMatrix(0, 1) = Format(grdTotal.ValueMatrix(0, 1) + grdMain.ValueMatrix(i, 1), "0.00")
        grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(i, 2), "0.00")
        grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(i, 3), "0.00")
        grdTotal.TextMatrix(0, 4) = Format(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(i, 4), "0.00")
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(i, 5), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(i, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(i, 7), "0.00")
        grdTotal.TextMatrix(0, 8) = grdTotal.ValueMatrix(0, 8) + Val(Replace(grdMain.TextMatrix(i, 8), "%", ""))
        grdTotal.TextMatrix(0, 9) = Format(grdTotal.ValueMatrix(0, 9) + grdMain.ValueMatrix(i, 9), "0.00")
    Next i
    grdTotal.TextMatrix(0, 8) = "0.00"
    grdTotal.TextMatrix(0, 0) = "Total Users: " & grdMain.Rows - 1
    rs.Close
    On Error GoTo 0
End Sub
Private Sub Sales_Analysis_by_Hour()
    Screen.MousePointer = 11
    grdTotal.Visible = False
    grdMain.Height = 8170
    ActiveReadServer "Select Dept_Name,Department_No from Departments where Dept_Type=1 order by Dept_Parent,convert(int,substring(Department_No,PATINDEX ( '%-%' , Department_No )+1,6))"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 25) = rs.Fields("Department_No")
        rs.MoveNext
    Wend
    rs.Close
    grdMain.Rows = grdMain.Rows + 1
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    For i = 1 To 24
        For b = 1 To grdMain.Rows - 1
            grdMain.TextMatrix(b, i) = "0.00"
            grdMain.Cell(flexcpBackColor, b, i, b, i) = vbWhite
        Next b
    Next i
    ActiveReadServer "Hour_Sales '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "','" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "','" & DeptString & "','" & LocString & "'"
    While Not rs.EOF
        Col = 0
        Select Case rs.Fields("Hour")
            Case 0: Col = 18
            Case 1: Col = 19
            Case 2: Col = 20
            Case 3: Col = 21
            Case 4: Col = 22
            Case 5: Col = 23
            Case 6: Col = 24
            Case 7: Col = 1
            Case 8: Col = 2
            Case 9: Col = 3
            Case 10: Col = 4
            Case 11: Col = 5
            Case 12: Col = 6
            Case 13: Col = 7
            Case 14: Col = 8
            Case 15: Col = 9
            Case 16: Col = 10
            Case 17: Col = 11
            Case 18: Col = 12
            Case 19: Col = 13
            Case 20: Col = 14
            Case 21: Col = 15
            Case 22: Col = 16
            Case 23: Col = 17
        End Select
        Row = grdMain.FindRow(rs.Fields("Department_no"), 0, 25)
        grdMain.TextMatrix(Row, Col) = Format(rs.Fields("Line_Total"), "0.00")
        grdMain.Cell(flexcpBackColor, Row, Col, Row, Col) = &HC0FFC0
        rs.MoveNext
    Wend
    rs.Close
    grdMain.TextMatrix(grdMain.Rows - 1, 0) = "0.00"
    For b = 1 To 24
        grdMain.TextMatrix(grdMain.Rows - 1, b) = "0.00"
    Next b
    For b = 1 To 24
        For i = 1 To grdMain.Rows - 2
            grdMain.TextMatrix(grdMain.Rows - 1, b) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, b) + grdMain.ValueMatrix(i, b), "0.00")
        Next i
    Next b
    For b = 1 To 24
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = grdMain.ValueMatrix(grdMain.Rows - 1, 0) + grdMain.ValueMatrix(grdMain.Rows - 1, b)
    Next b
    grdMain.TextMatrix(grdMain.Rows - 1, 0) = "Total Sales: " & Format(grdMain.TextMatrix(grdMain.Rows - 1, 0), "0.00")
    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 24) = &HE0E0E0
    grdMain.Cell(flexcpFontBold, grdMain.Rows - 1, 0, grdMain.Rows - 1, 24) = True
    Screen.MousePointer = 0
End Sub
Private Sub Placed_Purchase_Orders()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    If cmb1.Text <> "<All Suppliers>" Then
        If cmb3.Text = "<All Users>" Then
            ActiveReadServer "Select * from  Placed_Order_Listing Where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and Account_No ='" & Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1)) & "' order by Date_Time"
        Else
            ActiveReadServer "Select * from  Placed_Order_Listing Where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and Account_No ='" & Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1)) & "' and User_No = " & Val(Mid(cmb3.Text, 1, InStr(cmb3.Text, "-") - 1)) & " order by Date_Time"
        End If
    Else
        If cmb3.Text = "<All Users>" Or cmb3.Text = "" Then
            ActiveReadServer "Select * from  Placed_Order_Listing Where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' order by Date_Time"
        Else
            ActiveReadServer "Select * from  Placed_Order_Listing Where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and User_No = " & Val(Mid(cmb3.Text, 1, InStr(cmb3.Text, "-") - 1)) & " order by Date_Time"
        End If
    End If
    grdTotal.TextMatrix(0, 0) = " Totals:"
    grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = "0.00"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = Trim(rs.Fields("Supplier_Name")) & " - " & Trim(rs.Fields("Supplier_No"))
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = "Order No: " & Format(rs.Fields("Order_No"), "000000")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(Val(rs.Fields("Line_Total") & ""), "0.00")
        grdTotal.TextMatrix(0, 4) = Format(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
    If grdMain.Rows > 1 Then
         grdMain.Row = 1
         grdMain.SetFocus
    End If
End Sub
Private Sub Debtor_List()
    grdMain.Rows = 1
    Select Case cmb1.Text
    Case "<All Debtors>"
        ActiveReadServer "Select *,Case Debt_Type  WHEN 0 THEN 'Debtor' WHEN 1 THEN 'Staff Account' WHEN 2 THEN 'Management Account' WHEN 3 THEN 'Travel Agent' WHEN 4 THEN 'Member' END As DType from Debtors order by Debtor_Name"
    Case "Debtor Accounts"
        ActiveReadServer "Select *,Case Debt_Type  WHEN 0 THEN 'Debtor' WHEN 1 THEN 'Staff Account' WHEN 2 THEN 'Management Account' WHEN 3 THEN 'Travel Agent' WHEN 4 THEN 'Member' END As DType from Debtors where Debt_Type = 0 order by Debtor_Name"
    Case "Staff Accounts"
        ActiveReadServer "Select *,Case Debt_Type  WHEN 0 THEN 'Debtor' WHEN 1 THEN 'Staff Account' WHEN 2 THEN 'Management Account' WHEN 3 THEN 'Travel Agent' WHEN 4 THEN 'Member' END As DType from Debtors where Debt_Type = 1 order by Debtor_Name"
    Case "Management Accounts"
        ActiveReadServer "Select *,Case Debt_Type  WHEN 0 THEN 'Debtor' WHEN 1 THEN 'Staff Account' WHEN 2 THEN 'Management Account' WHEN 3 THEN 'Travel Agent' WHEN 4 THEN 'Member' END As DType from Debtors where Debt_Type = 2 order by Debtor_Name"
    Case "Travel Agents"
        ActiveReadServer "Select *,Case Debt_Type  WHEN 0 THEN 'Debtor' WHEN 1 THEN 'Staff Account' WHEN 2 THEN 'Management Account' WHEN 3 THEN 'Travel Agent' WHEN 4 THEN 'Member' END As DType from Debtors where Debt_Type = 3 order by Debtor_Name"
    Case "Members"
        ActiveReadServer "Select *,Case Debt_Type  WHEN 0 THEN 'Debtor' WHEN 1 THEN 'Staff Account' WHEN 2 THEN 'Management Account' WHEN 3 THEN 'Travel Agent' WHEN 4 THEN 'Member' END As DType from Debtors where Debt_Type = 4 order by Debtor_Name"
    End Select
    grdMain.Rows = rs.RecordCount + 1
    i = 0
    While Not rs.EOF
        i = i + 1
        grdMain.TextMatrix(i, 0) = rs.Fields("Debtor_No")
        grdMain.TextMatrix(i, 1) = rs.Fields("Debtor_Name")
        grdMain.TextMatrix(i, 2) = rs.Fields("Contact_Person") & ""
        grdMain.TextMatrix(i, 3) = rs.Fields("Business_Tel") & ""
        grdMain.TextMatrix(i, 4) = rs.Fields("DType") & ""
        grdMain.TextMatrix(i, 5) = Format(rs.Fields("Balance"), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    If grdMain.Rows > 1 Then grdMain.Row = 1
End Sub
Private Sub Dept_Age_Analysis()
    grdMain.Rows = 1
    Select Case cmb1.Text
        Case "<All Debtors>"
            DebtType = "%"
        Case "Debtor Accounts"
            DebtType = "0"
        Case "Staff Accounts"
            DebtType = "1"
        Case "Management Accounts"
            DebtType = "2"
        Case "Travel Agents"
            DebtType = "3"
        Case "Members"
            DebtType = "4"
    End Select


'    ActiveReadServer "Select Debtor_No,Debtor_Name,Balance,Case Debt_Type  WHEN 0 THEN 'Debtor' WHEN 1 THEN 'Staff Account' WHEN 2 THEN 'Management Account' WHEN 3 THEN 'Travel Agent' WHEN 4 THEN 'Member' END As DType" & _
'    " ,(Select sum(debit) from Debtor_Accounts where" & _
'    " (datepart(m,Date_Time) = datepart(m,getdate()) and date_Time<getdate())and Debtor_No like Debtor_Accounts.Account_No)" & _
'    " as CurrentB" & _

ActiveReadServer "Select Debtor_No,Debtor_Name,Balance,Case Debt_Type  WHEN 0 THEN 'Debtor' WHEN 1 THEN 'Staff Account' WHEN 2 THEN 'Management Account' WHEN 3 THEN 'Travel Agent' WHEN 4 THEN 'Member' END As DType" & _
    " ,(Select sum(debit) from Debtor_Accounts where" & _
    " (DATEDIFF(dd, Date_Time, GETDATE()) <=30) AND Debtor_No like Debtor_Accounts.Account_No)" & _
    " as CurrentB" & _
    " ,(Select sum(debit) from Debtor_Accounts where" & _
    " (DATEDIFF(dd, Date_Time, GETDATE()) <=60) AND (DATEDIFF(dd, Date_Time, GETDATE()) > 30) and Debtor_No like Debtor_Accounts.Account_No)" & _
    " as [30Days]" & _
    " ,(Select sum(debit) from Debtor_Accounts where" & _
    " (DATEDIFF(dd, Date_Time, GETDATE()) <= 90) AND (DATEDIFF(dd, Date_Time, GETDATE()) > 60) and Debtor_No like Debtor_Accounts.Account_No)" & _
    " as [60Days]" & _
    " ,(Select sum(debit) from Debtor_Accounts where" & _
    " (DATEDIFF(dd, Date_Time, GETDATE()) < =120) AND (DATEDIFF(dd, Date_Time, GETDATE()) >  90)and Debtor_No like Debtor_Accounts.Account_No)" & _
    " as [90Days]" & _
    " ,(Select sum(debit) from Debtor_Accounts where" & _
    " (DATEDIFF(dd, Date_Time, GETDATE()) >  120) and Debtor_No like Debtor_Accounts.Account_No)" & _
    " as [120Days+]" & _
    " from Debtors where Debt_Type like '" & DebtType & "' order by Debtor_Name"

    grdMain.Rows = rs.RecordCount + 1
    i = 0
    While Not rs.EOF
        i = i + 1
        grdMain.TextMatrix(i, 0) = rs.Fields("Debtor_No")
        grdMain.TextMatrix(i, 1) = rs.Fields("Debtor_Name")
        grdMain.TextMatrix(i, 2) = Format(Val(rs.Fields("120Days+") & ""), "0.00")
        grdMain.TextMatrix(i, 3) = Format(Val(rs.Fields("90Days") & ""), "0.00")
        grdMain.TextMatrix(i, 4) = Format(Val(rs.Fields("60Days") & ""), "0.00")
        grdMain.TextMatrix(i, 5) = Format(Val(rs.Fields("30Days") & ""), "0.00")
        grdMain.TextMatrix(i, 6) = Format(Val(rs.Fields("CurrentB") & ""), "0.00")
        grdMain.TextMatrix(i, 7) = Format(Val(rs.Fields("Balance") & ""), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    For i = 1 To grdMain.Rows - 1
        For b = 5 To 2 Step -1
            If grdMain.ValueMatrix(i, b) < 0 Then
                grdMain.TextMatrix(i, b - 1) = Format(grdMain.ValueMatrix(i, b) + grdMain.ValueMatrix(i, b - 1), "0.00")
                grdMain.TextMatrix(i, b) = "0.00"
            End If
        Next b
    Next i
    For i = 1 To grdMain.Rows - 1
        ActiveReadServer "Select sum(Credit) as Credit from Debtor_Accounts where Account_No = '" & grdMain.TextMatrix(i, 0) & "'"
        If rs.RecordCount > 0 Then
            Total_Credit = Val(rs.Fields("Credit") & "")
            For b = 2 To 6
                Total_Credit = Total_Credit - grdMain.ValueMatrix(i, b)
                If Total_Credit > 0 Then
                    grdMain.TextMatrix(i, b) = "0.00"
                Else
                    grdMain.TextMatrix(i, b) = Format(Abs(Total_Credit), "0.00")
                    Exit For
                End If
            Next b
        End If
        rs.Close
    Next i
top:
    Select Case cmb3.Text
        Case "Current"
            For i = 1 To grdMain.Rows - 1
                If grdMain.TextMatrix(i, 6) = "0.00" Then
                    grdMain.RemoveItem (i)
                    GoTo top
                End If
            Next i
        Case "30 Days"
            For i = 1 To grdMain.Rows - 1
                If grdMain.TextMatrix(i, 5) = "0.00" Then
                    grdMain.RemoveItem (i)
                    GoTo top
                End If
            Next i
        Case "60 Days"
            For i = 1 To grdMain.Rows - 1
                If grdMain.TextMatrix(i, 4) = "0.00" Then
                    grdMain.RemoveItem (i)
                    GoTo top
                End If
            Next i
        Case "90 Days"
            For i = 1 To grdMain.Rows - 1
                If grdMain.TextMatrix(i, 3) = "0.00" Then
                    grdMain.RemoveItem (i)
                    GoTo top
                End If
            Next i
        Case "120 Days+"
            For i = 1 To grdMain.Rows - 1
                If grdMain.TextMatrix(i, 2) = "0.00" Then
                    grdMain.RemoveItem (i)
                    GoTo top
                End If
            Next i
    End Select
    grdTotal.TextMatrix(0, 2) = "0.00"
    grdTotal.TextMatrix(0, 3) = "0.00"
    grdTotal.TextMatrix(0, 4) = "0.00"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    For i = 1 To grdMain.Rows - 1
        grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(i, 2), "0.00")
        grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(i, 3), "0.00")
        grdTotal.TextMatrix(0, 4) = Format(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(i, 4), "0.00")
        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(i, 5), "0.00")
        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(i, 6), "0.00")
        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(i, 7), "0.00")
    Next i
    If grdMain.Rows > 1 Then grdMain.Row = 1
End Sub
Private Sub Cost_Price_Var()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    DepString = "%"
    Suppstring = "%"
    If cmb1.Text = "<All Suppliers>" Then
        Suppstring = "%"
    Else
        If InStr(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2), "-") = 0 Then
            Suppstring = Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1)) & "%"
        Else
            Suppstring = Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1))
        End If
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    
    ActiveReadServer "Select *,(Select Supplier_Name from Suppliers where Suppliers.Supplier_No = Cost_Change_View.Supplier_No) as Supplier from Cost_Change_View " & _
    " where Department_No like '" & DeptString & "' and Supplier_No like '" & Suppstring & "' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Description"
    grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 1) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = "0.00"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    Ave_Tot = 0
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Supplier") & " - (" & Trim(rs.Fields("Supplier_No")) & ")"
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("Invoice_No")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Prev_Cost"), "0.00")
        If Val(rs.Fields("Prev_Cost") & "") = 0 Then
            grdMain.TextMatrix(grdMain.Rows - 1, 5) = "100%"
        Else
            grdMain.TextMatrix(grdMain.Rows - 1, 5) = Round(((Val(rs.Fields("Landed_Cost") & "") - Val(rs.Fields("Prev_Cost") & "")) / Val(rs.Fields("Prev_Cost") & "")) * 100, 2) & "%"
            Ave_Tot = Ave_Tot + grdMain.ValueMatrix(grdMain.Rows - 1, 5)
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Landed_Cost"), "0.00")
        rs.MoveNext
    Wend
    If rs.RecordCount - 1 <> 0 Then
        grdTotal.TextMatrix(0, 4) = "Average Price Change = " & Round(Ave_Tot / rs.RecordCount - 1, 2)
        grdTotal.TextMatrix(0, 5) = "Average Price Change = " & Round(Ave_Tot / rs.RecordCount - 1, 2)
        grdTotal.TextMatrix(0, 6) = "Average Price Change = " & Round(Ave_Tot / rs.RecordCount - 1, 2)
    Else
        grdTotal.TextMatrix(0, 4) = "Average Price Change = 0"
        grdTotal.TextMatrix(0, 5) = "Average Price Change = 0"
        grdTotal.TextMatrix(0, 6) = "Average Price Change = 0"
    End If
    rs.Close
    If grdMain.Rows > 1 Then
        grdMain.Row = 1
        grdMain.SetFocus
    End If
End Sub
Private Sub Table_Tranfer_Jour()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    If cmb1.Text <> "<All Users>" Then
        If cmb3.Text = "<All Users>" Then
            ActiveReadServer "Select * from  TabTable_View Where (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') and Tranfering_User ='" & Trim(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 1)) & "' order by Date_Time"
        Else
            ActiveReadServer "Select * from  TabTable_View Where (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') and Tranfering_User  ='" & Trim(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 1)) & "' and Recieving_User = " & Val(Mid(cmb3.Text, 1, InStr(cmb3.Text, "-") - 1)) & " order by Date_Time"
        End If
    Else
        If cmb3.Text = "<All Users>" Or cmb3.Text = "" Then
            ActiveReadServer "Select * from  TabTable_View Where (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Date_Time"
        Else
            ActiveReadServer "Select * from  TabTable_View Where (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') and Recieving_User = " & Val(Mid(cmb3.Text, 1, InStr(cmb3.Text, "-") - 1)) & " order by Date_Time"
        End If
    End If
    grdTotal.TextMatrix(0, 0) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 5) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 6) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 7) = " Transactions = " & rs.RecordCount
    grdTotal.TextMatrix(0, 8) = " Transactions = " & rs.RecordCount
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time") & "", "YYYY-MM-DD hh:mm")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Trans_User") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Rec_User") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("From_Table") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("To_Table") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = rs.Fields("From_Tab") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = rs.Fields("To_Tab") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = rs.Fields("Invoice_No") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = rs.Fields("User_Action") & ""
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub
Public Sub Pack_Links()
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    grdTotal.TextMatrix(0, 0) = ""
    grdTotal.TextMatrix(0, 1) = ""
    grdTotal.TextMatrix(0, 2) = ""
    grdTotal.TextMatrix(0, 3) = ""
    grdTotal.TextMatrix(0, 4) = ""
    grdTotal.TextMatrix(0, 5) = ""
    grdTotal.TextMatrix(0, 6) = ""
    ActiveReadServer "Select * from Products_Links_View where " & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'"
    grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 1) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 5) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 6) = " Products = " & rs.RecordCount
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Department") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("Pack_Size") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Ave_Cost") & "", "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = rs.Fields("Link_Code") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = rs.Fields("Link_Description")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub

Public Sub Selection_Change()
    cmb1.Visible = True
    On Error Resume Next
    DoEvents
    picLive.Visible = False
    If grdMain.Tag = "1" Then Exit Sub
    If cmb1.Text = "" Then Exit Sub
    grdMain.Rows = 1
    Screen.MousePointer = 11
    Select Case cmb2.Text
        Case "Pack Links"
            Pack_Links
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Table and Tab Journal"
            Table_Tranfer_Jour
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Cost Price Variations"
            Cost_Price_Var
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Age Analysis"
            Dept_Age_Analysis
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Debtor Accounts"
            Debtor_List
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Placed Purchase Orders"
            Placed_Purchase_Orders
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Analysis by Hour"
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
            Sales_Analysis_by_Hour
        Case "Staff Commision Report"
            User_Commission_List
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Payout Journal"
            Payout_Journal
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Pre-Sales Analysis by Price"
            Pre_Sales_Analysis
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Stock on Hand"
            Stock_on_Hand
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Stock on Hand (Suppliers)"
            Stock_on_Hand_Suppliers
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Stock Summary"
            Stock_Summary
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
    
        Case "Stock Levels Low"
            Stock_Levels
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
    
    
        Case "Product Analysis"
            Product_Analysis
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Analysis by Debtor"
            Sales_Analysis_Debtor
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Analysis by Supplier"
            Sales_Analysis_Supplier
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Stock Takes"
            Stock_Takes
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Corrections", "Discounted Sales"
            Sales_Correct
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Receive on Account"
            Receive_on_Account
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Staff Shift Report"
            User_Shift_List
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Deposits Paid", "Payments Received"
            Deposits_Paid
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Room Accounts"
            Room_Accounts
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Daily Trade Analysis"
            Day_trade
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
        
      Case "Sales Analysis by User type"
            Screen.MousePointer = 11
             
            frmMain.Toolbar1.Buttons(16).Enabled = False
            frmMain.Toolbar1.Buttons(16).Tag = ""
            cmb3.Left = 51730
            cmb1.Left = 51730
            Sabutreportgetready
        Screen.MousePointer = 1
        
       Case "Trade Comparison"
         Screen.MousePointer = 11
             
            frmMain.Toolbar1.Buttons(16).Enabled = False
            frmMain.Toolbar1.Buttons(16).Tag = ""
        TradeComparison
        Screen.MousePointer = 1
        Case "Transfer Journal"
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            If cmb1.Text <> "<Transfering Locations>" Then
                LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
            Else
                LocString = "%"
            End If
            If cmb3.Text <> "<Receiving Locations>" Then
                LocString1 = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
            Else
                LocString1 = "%"
            End If
            ActiveReadServer "Select * from Transfer_Journal_View where " & _
            " Trans_Location like '" & LocString & "'" & _
            " and Rec_Location like '" & LocString1 & "'" & _
            " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
            grdTotal.TextMatrix(0, 0) = ""
            grdTotal.TextMatrix(0, 3) = ""
            grdTotal.TextMatrix(0, 4) = ""
            grdTotal.TextMatrix(0, 5) = ""
            grdTotal.TextMatrix(0, 6) = "0"
            grdTotal.TextMatrix(0, 1) = " Transfers = " & rs.RecordCount
            grdTotal.TextMatrix(0, 2) = " Transfers = " & rs.RecordCount
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Date_Time")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("User") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Transfer_No") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = "Stock Transfer"
                grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Trans_Location")
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = rs.Fields("Rec_Location")
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Line_Total"), "0.00")
                grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
                rs.MoveNext
            Wend
            rs.Close
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Stock Variance"
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            If cmb1.Text <> "<All Locations>" Then
                LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
            Else
                LocString = "%"
            End If
            If cmb3.Text = "<All Departments>" Then
                DeptString = "%"
            Else
                If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
                    DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
                Else
                    DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
                End If
            End If
            grdMain.Rows = 1
            grdTotal.TextMatrix(0, 2) = "0"
            grdTotal.TextMatrix(0, 3) = "0.00"
            grdTotal.TextMatrix(0, 4) = "0"
            grdTotal.TextMatrix(0, 5) = "0.00"
            grdTotal.TextMatrix(0, 6) = "0.00"
            grdTotal.TextMatrix(0, 7) = "0.00"
            If cmb1.Text <> "<All Locations>" Then
                ActiveReadServer "Select * from Stock_Take_Variance where " & _
                " Department_No like '" & DeptString & "'" & _
                " and Location_No like '" & LocString & "'" & _
                " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
            Else
                ActiveReadServer "Select * from Stock_Var_Group where  Department_No like '" & DeptString & "'" & _
                " and (convert(DateTime,Date_Time) = '" & mthViewStart.Value & "')"
            End If
            grdTotal.TextMatrix(0, 0) = ""
            grdTotal.TextMatrix(0, 1) = " Products = " & rs.RecordCount
            grdTotal.TextMatrix(0, 2) = " Products = " & rs.RecordCount
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_code")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Department") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = Round(rs.Fields("Qty_on_Hand"), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(rs.Fields("Variance"), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = Round(rs.Fields("Qty_Counted"), 3)
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Variance") * rs.Fields("Ave_Cost"), "0.00")
                grdMain.TextMatrix(grdMain.Rows - 1, 7) = Format(rs.Fields("Qty_Counted") * rs.Fields("Ave_Cost"), "0.00")
                grdMain.TextMatrix(grdMain.Rows - 1, 8) = rs.Fields("Unit_Size")
                grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + Val(Replace(grdMain.TextMatrix(grdMain.Rows - 1, 4), "%", "")), 2)
                grdTotal.TextMatrix(0, 5) = Round(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), 3)
                grdTotal.TextMatrix(0, 3) = Round(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), 3)
                grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
                grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(grdMain.Rows - 1, 7), "0.00")
                rs.MoveNext
            Wend
            rs.Close
            For i = 1 To grdMain.Rows - 1
                Select Case grdMain.ValueMatrix(i, 8)
                    Case 750
                        Negative = False
                        If grdMain.ValueMatrix(i, 3) < 0 Then
                            Negative = True
                            grdMain.TextMatrix(i, 3) = Abs(grdMain.ValueMatrix(i, 3))
                        End If
                        Tots = Round((grdMain.ValueMatrix(i, 3) - Int(grdMain.ValueMatrix(i, 3))) * 30, 0)
                        If Round(grdMain.ValueMatrix(i, 3), 0) = 0 Then
                            If Tots = 0 Then
                                grdMain.TextMatrix(i, 3) = 0
                            Else
                                grdMain.TextMatrix(i, 3) = Tots & " (Tots)"
                            End If
                        Else
                            Tots = Round((grdMain.ValueMatrix(i, 3) - Int(grdMain.ValueMatrix(i, 3))) * 30, 0)
                            If Tots <> 0 Then
                                grdMain.TextMatrix(i, 3) = Int(grdMain.ValueMatrix(i, 3)) & " & " & Tots & " (Tots)"
                            End If
                        End If
                        If Negative = True Then grdMain.TextMatrix(i, 3) = "-" & grdMain.TextMatrix(i, 3)
                        
                        Negative = False
                        If grdMain.ValueMatrix(i, 4) < 0 Then
                            Negative = True
                            grdMain.TextMatrix(i, 4) = Abs(grdMain.ValueMatrix(i, 4))
                        End If
                        Tots = Round((grdMain.ValueMatrix(i, 4) - Int(grdMain.ValueMatrix(i, 4))) * 30, 0)
                        If Round(grdMain.ValueMatrix(i, 4), 0) = 0 Then
                            If Tots = 0 Then
                                grdMain.TextMatrix(i, 4) = 0
                            Else
                                grdMain.TextMatrix(i, 4) = Tots & " (Tots)"
                            End If
                        Else
                            Tots = Round((grdMain.ValueMatrix(i, 4) - Int(grdMain.ValueMatrix(i, 4))) * 30, 0)
                            If Tots <> 0 Then
                                grdMain.TextMatrix(i, 4) = Int(grdMain.ValueMatrix(i, 4)) & " & " & Tots & " (Tots)"
                            End If
                        End If
                        If Negative = True Then grdMain.TextMatrix(i, 4) = "-" & grdMain.TextMatrix(i, 4)
                        
                        Negative = False
                        If grdMain.ValueMatrix(i, 5) < 0 Then
                            Negative = True
                            grdMain.TextMatrix(i, 5) = Abs(grdMain.ValueMatrix(i, 5))
                        End If
                        Tots = Round((grdMain.ValueMatrix(i, 5) - Int(grdMain.ValueMatrix(i, 5))) * 30, 0)
                        If Round(grdMain.ValueMatrix(i, 5), 0) = 0 Then
                            If Tots = 0 Then
                                grdMain.TextMatrix(i, 5) = 0
                            Else
                                grdMain.TextMatrix(i, 5) = Tots & " (Tots)"
                            End If
                        Else
                            Tots = Round((grdMain.ValueMatrix(i, 5) - Int(grdMain.ValueMatrix(i, 5))) * 30, 0)
                            If Tots <> 0 Then
                                grdMain.TextMatrix(i, 5) = Int(grdMain.ValueMatrix(i, 5)) & " & " & Tots & " (Tots)"
                            End If
                        End If
                        If Negative = True Then grdMain.TextMatrix(i, 5) = "-" & grdMain.TextMatrix(i, 5)
                End Select
            Next i
            grdTotal.MergeRow(0) = True
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Stock Movement (Quantities)"
            Stock_Movement_Quantities
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Stock Movement (Values)"
            Stock_Movement_Values
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "User Journals"
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            If cmb1.Text <> "<All Users>" Then
                ActiveReadServer "Select * from User_Journal_View where User_No= " & Trim(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & " and Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "' order by Date_Time"
            Else
                ActiveReadServer "Select * from User_Journal_View where Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "' order by Date_Time"
            End If
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY HH:mm:ss AM/PM")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("User_No") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Workstation_No") & " - " & rs.Fields("Workstation_Name")
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("Key_Desc")
                If rs.Fields("Function_Key") = 3 Then
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 3) = &HD7FDD9
                End If
                If rs.Fields("Function_Key") = 4 Then
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 3) = &HDBDEFD
                End If
                If rs.Fields("Function_Key") = 5 Then
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 3) = &HC0FFFF
                End If
                rs.MoveNext
            Wend
            grdTotal.TextMatrix(0, 0) = " Totals:"
            grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
            grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
            grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
            rs.Close
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        
        Case "Purchase Journal"
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            If cmb3.Text = "<All Departments>" Then
                DeptString = "%"
            Else
                If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
                    DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
                Else
                    DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
                End If
            End If
            FunctionKey = 16
            If cmb1.Text <> "<All Suppliers>" Then
                If cmb3.Text = "<All Departments>" Then
                    ActiveReadServer "Select * from Purchase_Journal_View where Supplier_No= '" & Trim(Mid(cmb1.Text, InStr(cmb1.Text, "-") + 1)) & "' and Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' order by Date_Time"
                Else
                    ActiveReadServer "Select * from Purchase_JournalDept_View where Department_No like '" & DeptString & "' and  Suppliers_No= '" & Trim(Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 1)) & "' and  Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and Function_Key = " & FunctionKey & " order by Date_Time"
                End If
            Else
                If cmb3.Text = "<All Departments>" Or cmb3.Text = "" Then
                    ActiveReadServer "Select * from Purchase_Journal_View where Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' order by Date_Time"
                Else
                    ActiveReadServer "Select * from Purchase_JournalDept_View where  Department_No like '" & DeptString & "' and Date_Time > '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' and Function_Key = " & FunctionKey & " order by Date_Time"
                End If
            End If
            grdTotal.TextMatrix(0, 0) = " Totals:"
            grdTotal.TextMatrix(0, 1) = " Transactions = " & rs.RecordCount
            grdTotal.TextMatrix(0, 2) = " Transactions = " & rs.RecordCount
            grdTotal.TextMatrix(0, 3) = " Transactions = " & rs.RecordCount
            grdTotal.TextMatrix(0, 4) = " Transactions = " & rs.RecordCount
            grdTotal.TextMatrix(0, 5) = "0"
            grdTotal.TextMatrix(0, 6) = "0"
            grdTotal.TextMatrix(0, 7) = "0"
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "DD MMM YYYY")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("User_No") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = Trim(rs.Fields("Supplier_Name")) & " - " & Trim(rs.Fields("Supplier_No"))
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Invoice_No"), "00000")
                grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Key_Desc") & ""
                If grdMain.TextMatrix(grdMain.Rows - 1, 4) = "Goods Received Voucher" Then
                    Select Case rs.Fields("Debit")
                        Case Is > 0
                            grdMain.TextMatrix(grdMain.Rows - 1, 4) = "GRV - " & rs.Fields("Tender_Type") & " > Ref: " & rs.Fields("Ref_No") & " R " & Format(rs.Fields("Debit"), "0.00")
                        Case Else
                            grdMain.TextMatrix(grdMain.Rows - 1, 4) = "GRV - Not Paid"
                    End Select
                End If
                grdMain.TextMatrix(grdMain.Rows - 1, 7) = Format(rs.Fields("Line_Total") + rs.Fields("Vat_Rate"), "0.00")
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Vat_Rate"), "0.00")
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total"), "0.00")
                grdMain.TextMatrix(grdMain.Rows - 1, 8) = rs.Fields("GRV_No")
                grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
                grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
                grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(grdMain.Rows - 1, 7), "0.00")
                If rs.Fields("Function_Key") = 6 Then
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 7) = &HD7FDD9
                End If
                rs.MoveNext
            Wend
            rs.Close
            grdTotal.MergeRow(0) = True
            If grdMain.Rows > 1 Then
                 grdMain.Row = 1
                 grdMain.SetFocus
            End If
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Room Sales"
            Room_Sales
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Journal"
            Sales_Journal
     frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Analysis by Product"
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            ActiveReadServer "SELECT Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product') AS Description," & _
            "SUM(Sales_Journal.Sales_Tax) AS Sales_Tax, SUM(Sales_Journal.Line_Total) AS Line_Total, Sales_Journal.Department_No," & _
            "SUM(Sales_Journal.Qty) AS Qty, SUM(Sales_Journal.Ave_Cost*Sales_Journal.Qty) AS Ave_Cost," & _
            "SUM(Sales_Journal.Line_Total / ((100 + Sales_Journal.Sales_Tax) / 100)) As Line_Total_Excl " & _
            "FROM Sales_Journal LEFT OUTER JOIN " & _
            "Products ON Sales_Journal.Product_Code = Products.Product_Code " & _
            "WHERE (Sales_Journal.Ave_Cost + Sales_Journal.Line_Total <> 0) and (Sales_Journal.Function_Key IN (7)) AND ((ISNULL(Sales_Journal.Extra, '') = '' or Sales_Journal.Extra='Return Item')) and " & _
            "(Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
            " GROUP BY Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product'), Sales_Journal.Department_No " & _
            "HAVING (Sales_Journal.Product_Code <> '')"
            grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
            grdTotal.TextMatrix(0, 2) = "0"
            grdTotal.TextMatrix(0, 3) = "0.00"
            grdTotal.TextMatrix(0, 4) = "0"
            grdTotal.TextMatrix(0, 5) = "0.00"
            grdTotal.TextMatrix(0, 6) = "0.00"
            grdTotal.TextMatrix(0, 7) = "0.00"
            ToTGPValue = 0
            TotTaxValue1 = 0
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Qty")
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Ave_Cost"), "0.00")
                If Round(rs.Fields("Line_Total_Excl"), 3) = 0 Then
                    grdMain.TextMatrix(grdMain.Rows - 1, 4) = "0%"
                Else
                    
                    grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(((rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")) / rs.Fields("Line_Total_Excl")) * 100, 2) & "%"
                End If
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Line_Total"), "0.00")
                TotTaxValue1 = TotTaxValue1 + (rs.Fields("Line_Total") - rs.Fields("Line_Total_Excl"))
                ToTGPValue = ToTGPValue + rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost"), "0.00")
                grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), "0.00")
                grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + Val(Replace(grdMain.TextMatrix(grdMain.Rows - 1, 4), "%", "")), 2)
                grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
                grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
                grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
                grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 6), "0.00")
                rs.MoveNext
            Wend
            If rs.RecordCount = 0 Then
                grdTotal.TextMatrix(0, 4) = "0%"
            Else
                grdTotal.TextMatrix(0, 4) = Round((ToTGPValue / (grdTotal.ValueMatrix(0, 6) - TotTaxValue1)) * 100, 3) & " %"
            End If
            rs.Close
            If grdTotal.ValueMatrix(0, 2) = 0 Then
                AveGP = 0
            Else
                AveGP = ToTGPValue / grdTotal.ValueMatrix(0, 2)
            End If
            grdTotal.TextMatrix(0, 1) = " Ave GP Value: " & Format(AveGP, "0.00")
            grdTotal.MergeRow(0) = True
            newAve = 0
            RowCount = 0
            For i = 1 To grdMain.Rows - 1
                grdMain.RowHidden(i) = False
                If grdMain.ValueMatrix(i, 2) = 0 Then
                    grdMain.TextMatrix(i, 7) = "Problem Child"
                Else
                    grdMain.TextMatrix(i, 7) = Round(((grdMain.ValueMatrix(i, 5) / grdMain.ValueMatrix(i, 2)) / AveGP) * 100, 2)
                    Select Case grdMain.ValueMatrix(i, 7)
                        Case Is >= 100
                            grdMain.TextMatrix(i, 7) = "Star"
                        Case Is > 66.66
                            grdMain.TextMatrix(i, 7) = "Cash Cow"
                        Case Is > 33.33
                            grdMain.TextMatrix(i, 7) = "Dog"
                        Case Is > 0
                            grdMain.TextMatrix(i, 7) = "Problem Child"
                        Case 0
                            grdMain.TextMatrix(i, 7) = "Star"
                        Case Is < 0
                            grdMain.TextMatrix(i, 7) = "Star"
                    End Select
                End If

                If cmb1.Text <> "<All Products>" Then
                    If cmb1.Text <> grdMain.TextMatrix(i, 7) Then
                        grdMain.RowHidden(i) = True
                        grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) - grdMain.ValueMatrix(i, 2), "0.00")
                        grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) - grdMain.ValueMatrix(i, 5), "0.00")
                        grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) - grdMain.ValueMatrix(i, 3), "0.00")
                        grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) - grdMain.ValueMatrix(i, 6), "0.00")
                        grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 6), "0.00")
                    Else
                        newAve = Format(newAve + Val(Replace(grdMain.TextMatrix(i, 4), "%", "")), "0.00")
                        RowCount = RowCount + 1
                    End If
                End If
            Next i
            If cmb1.Text <> "<All Products>" Then
                If RowCount > 0 Then
                    grdTotal.TextMatrix(0, 4) = Round(newAve / RowCount, 2) & "%"
                    grdTotal.TextMatrix(0, 0) = " Products = " & RowCount
                End If
            End If
            If grdMain.Rows > 1 Then
                grdMain.Row = 1
                grdMain.SetFocus
            End If
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Analysis by Department"
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            If cmb1.Text = "<All Departments>" Then
                DeptString = "%"
            Else
                If InStr(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2), "-") = 0 Then
                    DeptString = Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2) & "%"
                Else
                    DeptString = Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2)
                End If
            End If
            ActiveReadServer "SELECT Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product') AS Description," & _
            "SUM(Sales_Journal.Sales_Tax) AS Sales_Tax, SUM(Sales_Journal.Line_Total) AS Line_Total, Sales_Journal.Department_No," & _
            "SUM(Sales_Journal.Qty) AS Qty, SUM(Sales_Journal.Ave_Cost*Sales_Journal.Qty) AS Ave_Cost," & _
            "SUM(Sales_Journal.Line_Total / ((100 + Sales_Journal.Sales_Tax) / 100)) As Line_Total_Excl " & _
            "FROM Sales_Journal LEFT OUTER JOIN " & _
            "Products ON Sales_Journal.Product_Code = Products.Product_Code " & _
            "WHERE (Sales_Journal.Ave_Cost + Sales_Journal.Line_Total <> 0) and (Sales_Journal.Function_Key IN (7))  and Sales_Journal.Department_No like '" & DeptString & "' AND ((ISNULL(Sales_Journal.Extra, '') = '' or Sales_Journal.Extra='Return Item')) and " & _
            "(Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
            " GROUP BY Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product'), Sales_Journal.Department_No " & _
            "HAVING (Sales_Journal.Product_Code <> '')"
            grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
            grdTotal.TextMatrix(0, 2) = "0"
            grdTotal.TextMatrix(0, 3) = "0.00"
            grdTotal.TextMatrix(0, 4) = "0"
            grdTotal.TextMatrix(0, 5) = "0.00"
            grdTotal.TextMatrix(0, 6) = "0.00"
            grdTotal.TextMatrix(0, 7) = "0.00"
            ToTGPValue = 0
            TotTaxValue1 = 0
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Qty")
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Ave_Cost"), "0.00")
                If Round(rs.Fields("Line_Total_Excl"), 3) = 0 Then
                    grdMain.TextMatrix(grdMain.Rows - 1, 4) = "0%"
                Else
                    grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(((rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")) / rs.Fields("Line_Total_Excl")) * 100, 2) & "%"
                End If
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Line_Total"), "0.00")
                TotTaxValue1 = TotTaxValue1 + (rs.Fields("Line_Total") - rs.Fields("Line_Total_Excl"))
                ToTGPValue = ToTGPValue + rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost"), "0.00")
                grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), "0.00")
                grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + Val(Replace(grdMain.TextMatrix(grdMain.Rows - 1, 4), "%", "")), 2)
                grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
                grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
                grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
                grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 6), "0.00")
                rs.MoveNext
            Wend
            If rs.RecordCount = 0 Then
                grdTotal.TextMatrix(0, 4) = "0%"
            Else
                grdTotal.TextMatrix(0, 4) = Round((ToTGPValue / (grdTotal.ValueMatrix(0, 6) - TotTaxValue1)) * 100, 3) & " %"
            End If
            rs.Close
            If grdTotal.ValueMatrix(0, 2) = 0 Then
                AveGP = 0
            Else
                AveGP = ToTGPValue / grdTotal.ValueMatrix(0, 2)
            End If
            grdTotal.TextMatrix(0, 1) = " Ave GP Value: " & Format(AveGP, "0.00")
            grdTotal.MergeRow(0) = True
            For i = 1 To grdMain.Rows - 1
                grdMain.RowHidden(i) = False
                If grdMain.ValueMatrix(i, 2) = 0 Then
                    grdMain.TextMatrix(i, 7) = "Problem Child"
                Else
                    grdMain.TextMatrix(i, 7) = Round(((grdMain.ValueMatrix(i, 5) / grdMain.ValueMatrix(i, 2)) / AveGP) * 100, 2)
                    Select Case grdMain.ValueMatrix(i, 7)
                        Case Is >= 100
                            grdMain.TextMatrix(i, 7) = "Star"
                        Case Is > 66.66
                            grdMain.TextMatrix(i, 7) = "Cash Cow"
                        Case Is > 33.33
                            grdMain.TextMatrix(i, 7) = "Dog"
                        Case Is > 0
                            grdMain.TextMatrix(i, 7) = "Problem Child"
                        Case 0
                            grdMain.TextMatrix(i, 7) = "Star"
                        Case Is < 0
                            grdMain.TextMatrix(i, 7) = "Star"
                    End Select
                End If
            Next i
            If grdMain.Rows > 1 Then
                grdMain.Row = 1
                grdMain.SetFocus
            End If
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Analysis by User"
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            If cmb1.Text = "<All Users>" Then
                useString = "%"
            Else
                If InStr(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2), "-") = 0 Then
                    useString = Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2) & "%"
                Else
                    useString = Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2)
                End If
            End If
            If cmb3.Text = "<All Departments>" Then
                DeptString = "%"
            Else
                If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
                    DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
                Else
                    DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
                End If
            End If
            ActiveReadServer "SELECT Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product') AS Description," & _
            "SUM(Sales_Journal.Sales_Tax) AS Sales_Tax, SUM(Sales_Journal.Line_Total) AS Line_Total, Sales_Journal.Department_No," & _
            "SUM(Sales_Journal.Qty) AS Qty, SUM(Sales_Journal.Ave_Cost*Sales_Journal.Qty) AS Ave_Cost," & _
            "SUM(Sales_Journal.Line_Total / ((100 + Sales_Journal.Sales_Tax) / 100)) As Line_Total_Excl " & _
            "FROM Sales_Journal LEFT OUTER JOIN " & _
            "Products ON Sales_Journal.Product_Code = Products.Product_Code " & _
            "WHERE (Sales_Journal.Ave_Cost + Sales_Journal.Line_Total <> 0) and (Sales_Journal.Function_Key IN (7))  and Sales_Journal.Department_No like '" & DeptString & "'AND User_No like '" & useString & "' AND ((ISNULL(Sales_Journal.Extra, '') = '' or Sales_Journal.Extra='Return Item')) and " & _
            "(Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
            " GROUP BY Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product'), Sales_Journal.Department_No " & _
            "HAVING (Sales_Journal.Product_Code <> '')"
            grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
            grdTotal.TextMatrix(0, 2) = "0"
            grdTotal.TextMatrix(0, 3) = "0.00"
            grdTotal.TextMatrix(0, 4) = "0"
            grdTotal.TextMatrix(0, 5) = "0.00"
            grdTotal.TextMatrix(0, 6) = "0.00"
            grdTotal.TextMatrix(0, 7) = "0.00"
            ToTGPValue = 0
            TotTaxValue1 = 0
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Qty")
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Ave_Cost"), "0.00")
                If Round(rs.Fields("Line_Total_Excl"), 3) = 0 Then
                    grdMain.TextMatrix(grdMain.Rows - 1, 4) = "0%"
                Else
                    grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(((rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")) / rs.Fields("Line_Total_Excl")) * 100, 2) & "%"
                End If
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Line_Total"), "0.00")
                ToTGPValue = ToTGPValue + rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")
                TotTaxValue1 = TotTaxValue1 + (rs.Fields("Line_Total") - rs.Fields("Line_Total_Excl"))
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost"), "0.00")
                grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), "0.00")
                grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + Val(Replace(grdMain.TextMatrix(grdMain.Rows - 1, 4), "%", "")), 2)
                grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
                grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
                grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
                grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 6), "0.00")
                rs.MoveNext
            Wend
            If rs.RecordCount = 0 Then
                grdTotal.TextMatrix(0, 4) = "0%"
            Else
                grdTotal.TextMatrix(0, 4) = Round((ToTGPValue / (grdTotal.ValueMatrix(0, 6) - TotTaxValue1)) * 100, 3) & " %"
            End If
            rs.Close
            If grdTotal.ValueMatrix(0, 2) = 0 Then
                AveGP = 0
            Else
                AveGP = ToTGPValue / grdTotal.ValueMatrix(0, 2)
            End If
            grdTotal.TextMatrix(0, 1) = " Ave GP Value: " & Format(AveGP, "0.00")
            grdTotal.MergeRow(0) = True
            For i = 1 To grdMain.Rows - 1
                grdMain.RowHidden(i) = False
                If grdMain.ValueMatrix(i, 2) = 0 Then
                    grdMain.TextMatrix(i, 7) = "Problem Child"
                Else
                    grdMain.TextMatrix(i, 7) = Round(((grdMain.ValueMatrix(i, 5) / grdMain.ValueMatrix(i, 2)) / AveGP) * 100, 2)
                    Select Case grdMain.ValueMatrix(i, 7)
                        Case Is >= 100
                            grdMain.TextMatrix(i, 7) = "Star"
                        Case Is > 66.66
                            grdMain.TextMatrix(i, 7) = "Cash Cow"
                        Case Is > 33.33
                            grdMain.TextMatrix(i, 7) = "Dog"
                        Case Is > 0
                            grdMain.TextMatrix(i, 7) = "Problem Child"
                        Case 0
                            grdMain.TextMatrix(i, 7) = "Star"
                        Case Is < 0
                            grdMain.TextMatrix(i, 7) = "Star"
                    End Select
                End If
            Next i
            If grdMain.Rows > 1 Then
                grdMain.Row = 1
                grdMain.SetFocus
            End If
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Sales Analysis by Location"
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            If cmb1.Text = "<All Locations>" Then
                LocString = "%"
            Else
                LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
            End If
            ActiveReadServer "SELECT Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product') AS Description," & _
            "SUM(Sales_Journal.Sales_Tax) AS Sales_Tax, SUM(Sales_Journal.Line_Total) AS Line_Total, Sales_Journal.Department_No," & _
            "SUM(Sales_Journal.Qty) AS Qty, SUM(Sales_Journal.Ave_Cost*Sales_Journal.Qty) AS Ave_Cost," & _
            "SUM(Sales_Journal.Line_Total / ((100 + Sales_Journal.Sales_Tax) / 100)) As Line_Total_Excl " & _
            "FROM Sales_Journal LEFT OUTER JOIN " & _
            "Products ON Sales_Journal.Product_Code = Products.Product_Code " & _
            "WHERE (Sales_Journal.Ave_Cost + Sales_Journal.Line_Total <> 0) and (Sales_Journal.Function_Key IN (7))  and Sales_Journal.Location like '" & LocString & "' AND ((ISNULL(Sales_Journal.Extra, '') = '' or Sales_Journal.Extra='Return Item')) and " & _
            "(Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
            " GROUP BY Sales_Journal.Product_Code, ISNULL(Products.Description, N'Deleted Product'), Sales_Journal.Department_No " & _
            "HAVING (Sales_Journal.Product_Code <> '')"
            grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
            grdTotal.TextMatrix(0, 2) = "0"
            grdTotal.TextMatrix(0, 3) = "0.00"
            grdTotal.TextMatrix(0, 4) = "0"
            grdTotal.TextMatrix(0, 5) = "0.00"
            grdTotal.TextMatrix(0, 6) = "0.00"
            grdTotal.TextMatrix(0, 7) = "0.00"
            ToTGPValue = 0
            TotTaxValue1 = 0
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Qty")
                grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("Ave_Cost"), "0.00")
                If Round(rs.Fields("Line_Total_Excl"), 3) = 0 Then
                    grdMain.TextMatrix(grdMain.Rows - 1, 4) = "0%"
                Else
                    grdMain.TextMatrix(grdMain.Rows - 1, 4) = Round(((rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")) / rs.Fields("Line_Total_Excl")) * 100, 2) & "%"
                End If
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Line_Total"), "0.00")
                TotTaxValue1 = TotTaxValue1 + (rs.Fields("Line_Total") - rs.Fields("Line_Total_Excl"))
                ToTGPValue = ToTGPValue + rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost")
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Line_Total_Excl") - rs.Fields("Ave_Cost"), "0.00")
                grdTotal.TextMatrix(0, 2) = Format(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), "0.00")
                grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + Val(Replace(grdMain.TextMatrix(grdMain.Rows - 1, 4), "%", "")), 2)
                grdTotal.TextMatrix(0, 5) = Format(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), "0.00")
                grdTotal.TextMatrix(0, 3) = Format(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
                grdTotal.TextMatrix(0, 6) = Format(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), "0.00")
                grdTotal.TextMatrix(0, 7) = Format(grdTotal.ValueMatrix(0, 6), "0.00")
                rs.MoveNext
            Wend
            If rs.RecordCount = 0 Then
                grdTotal.TextMatrix(0, 4) = "0%"
            Else
                grdTotal.TextMatrix(0, 4) = Round((ToTGPValue / (grdTotal.ValueMatrix(0, 6) - TotTaxValue1)) * 100, 3) & " %"
            End If
            rs.Close
            If grdTotal.ValueMatrix(0, 2) = 0 Then
                AveGP = 0
            Else
                AveGP = ToTGPValue / grdTotal.ValueMatrix(0, 2)
            End If
            grdTotal.TextMatrix(0, 1) = " Ave GP Value: " & Format(AveGP, "0.00")
            grdTotal.MergeRow(0) = True
            For i = 1 To grdMain.Rows - 1
                grdMain.RowHidden(i) = False
                If grdMain.ValueMatrix(i, 2) = 0 Then
                    grdMain.TextMatrix(i, 7) = "Problem Child"
                Else
                    grdMain.TextMatrix(i, 7) = Round(((grdMain.ValueMatrix(i, 5) / grdMain.ValueMatrix(i, 2)) / AveGP) * 100, 2)
                    Select Case grdMain.ValueMatrix(i, 7)
                        Case Is >= 100
                            grdMain.TextMatrix(i, 7) = "Star"
                        Case Is > 66.66
                            grdMain.TextMatrix(i, 7) = "Cash Cow"
                        Case Is > 33.33
                            grdMain.TextMatrix(i, 7) = "Dog"
                        Case Is > 0
                            grdMain.TextMatrix(i, 7) = "Problem Child"
                        Case 0
                            grdMain.TextMatrix(i, 7) = "Star"
                        Case Is < 0
                            grdMain.TextMatrix(i, 7) = "Star"
                    End Select
                End If
            Next i
            If grdMain.Rows > 1 Then
                grdMain.Row = 1
                grdMain.SetFocus
            End If
    frmMain.Toolbar1.Buttons(16).Enabled = True
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        Case "Trade Analysis"
            Trade_Analysis
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
    
   
    Case "Sales Analysis by User type"
      Sabutreportgetready
     
    Case "Trade Comparison"
    TradeComparison
        
    End Select
    Screen.MousePointer = 0
    On Error GoTo 0
End Sub
Private Sub Trade_Analysis()
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdTrans.TextMatrix(1, 1) = "0.00"
    grdTrans.TextMatrix(1, 2) = "0"
    grdTrans.TextMatrix(2, 1) = "0.00"
    grdTrans.TextMatrix(2, 2) = "0"
    grdTrans.TextMatrix(3, 1) = "0.00"
    grdTrans.TextMatrix(3, 2) = "0"
    grdTrans.TextMatrix(4, 1) = "0.00"
    grdTrans.TextMatrix(4, 2) = "0"
    grdTrans.TextMatrix(5, 2) = "0"
    grdTrans.TextMatrix(5, 1) = "0.00"
    
    grdRev.TextMatrix(1, 1) = "0.00"
    grdRev.TextMatrix(2, 1) = "0.00"
    grdRev.TextMatrix(3, 1) = "0.00"
    grdRev.TextMatrix(4, 1) = "0.00"
    grdRev.TextMatrix(5, 1) = "0.00"
    grdRev.TextMatrix(6, 1) = "0.00"
    grdRev.TextMatrix(1, 2) = "0"
    grdRev.TextMatrix(2, 2) = "0"
    grdRev.TextMatrix(3, 2) = "0"
    grdRev.TextMatrix(4, 2) = "0"
    grdRev.TextMatrix(5, 2) = "0"
    grdRev.TextMatrix(6, 2) = "0"
    grdTax.TextMatrix(1, 1) = "0.00"
    grdTax.TextMatrix(2, 1) = "0.00"
    grdTax.TextMatrix(3, 1) = "0.00"
    grdTax.TextMatrix(4, 1) = "0.00"
    grdTax.TextMatrix(5, 1) = "0.00"
    grdStock.TextMatrix(2, 1) = "0.00"
    factor = 1
    
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    

    
      ActiveReadServer "Select isnull(" & _
    " (SELECT SUM(Line_Total) AS Line_Total" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20))" & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " )/" & _
    " (SELECT SUM(Line_Total) AS Line_Total" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20))" & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '%'" & _
    " and Location like '%'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " ),0) as Factor"
    
 
    factor = rs.Fields("Factor")
    rs.Close
    ActiveReadServer "SELECT Function_Key," & _
        "Case Function_Key" & _
            " WHEN 9 THEN 'Cash'" & _
            " WHEN 10 THEN 'Card'" & _
            " WHEN 11 THEN 'Voucher'" & _
            " WHEN 12 THEN 'Charge'" & _
            " WHEN 13 THEN 'Loyalty'" & _
        " END As Sale_Type" & _
        ",COUNT(Line_No) AS Tend_Count,SUM(Sales_Tax) AS Sales_Tax,SUM(Line_Total) AS Line_Total,Sum(Ave_Cost) as Ave_Cost" & _
        " From Sales_Journal" & _
        " WHERE (Function_Key IN ( 9, 10, 11, 12, 13)) " & _
        " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
        " GROUP BY Function_Key"
    While Not rs.EOF
        Select Case rs.Fields("Sale_Type")
            Case "Cash"
                grdRev.TextMatrix(1, 1) = Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(1, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Card"
                grdRev.TextMatrix(2, 1) = Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(2, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Voucher"
                grdRev.TextMatrix(3, 1) = Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(3, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Charge"
                grdRev.TextMatrix(4, 1) = Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(4, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Loyalty"
                grdRev.TextMatrix(5, 1) = Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(5, 2) = Val(rs.Fields("Tend_Count") & "")
        End Select
        
        grdRev.TextMatrix(6, 1) = Format(grdRev.ValueMatrix(6, 1) + Round(Val(rs.Fields("Line_Total") * factor & ""), 2), "0.00")
        grdRev.TextMatrix(6, 2) = grdRev.ValueMatrix(6, 2) + Val(rs.Fields("Tend_Count") & "")
                        
        grdTax.TextMatrix(3, 1) = Format(grdTax.ValueMatrix(3, 1) + Round(Val(rs.Fields("Sales_Tax") & "") * factor, 2), "0.00")
        grdTax.TextMatrix(4, 1) = Format(grdTax.ValueMatrix(4, 1) + Round(Val(rs.Fields("Sales_Tax") & "") * factor, 2), "0.00")
        
        rs.MoveNext
    Wend
    rs.Close
    
    'Kotie
    'ActiveReadServer "Select isnull(SUM(Line_Total),0) AS Taxable" & _
    '" From Sales_Journal" & _
    '" Where (Function_Key in(7,20)) and Sales_Tax <> 0" & _
    '" and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    '" and Department_No like '" & DeptString & "'" & _
    '" and Location like '" & LocString & "'" & _
    '" and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"

    ActiveReadServer "Select isnull(SUM(Line_Total),0) AS Taxable" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20)) and Sales_Tax <> 0" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"

    grdTax.TextMatrix(1, 1) = Format(grdTax.ValueMatrix(1, 1) + Round(Val(rs.Fields("Taxable") & "") * factor, 2), "0.00")
    rs.Close
    
    ActiveReadServer "Select isnull(SUM(Line_Total),0) AS Taxable" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20)) and Sales_Tax = 0" & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"

    grdTax.TextMatrix(2, 1) = Format(grdTax.ValueMatrix(2, 1) + Round(Val(rs.Fields("Taxable") & "") * factor, 2), "0.00")
    rs.Close
    
    If Branch_Type = 5 And Selender > Date Then
        picLive.Visible = True
        ActiveReadServer "Select sum(isnull(line_Total,0)) as Total from Table_Listing where isnull(Extra_Function,'') = ''"
        If rs.Fields("Total") = 0 Then
            lblLive.Caption = "You have no Open Tables"
        Else
            lblLive.Caption = "Total with Open Tables: " & Format(rs.Fields("Total") + grdRev.ValueMatrix(6, 1), "0.00")
        End If
        rs.Close
    End If
    
    grdTrans.TextMatrix(5, 2) = "0"
    grdTrans.TextMatrix(5, 1) = "0.00"
    
    grdCount.TextMatrix(1, 1) = "0.00"
    grdCount.TextMatrix(2, 1) = "0.00"
    grdCount.TextMatrix(3, 1) = "0.00"
    grdCount.TextMatrix(4, 1) = "0.00"
    grdCount.TextMatrix(5, 1) = "0.00"
    grdCount.TextMatrix(6, 1) = "0.00"
    grdCount.TextMatrix(7, 1) = "0.00"
    grdCount.TextMatrix(1, 2) = "0"
    grdCount.TextMatrix(2, 2) = "0"
    grdCount.TextMatrix(3, 2) = "0"
    grdCount.TextMatrix(4, 2) = "0"
    grdCount.TextMatrix(5, 2) = "0"
    grdCount.TextMatrix(6, 2) = "0"
    grdCount.TextMatrix(7, 2) = "0"
    
    ActiveReadServer "Select sum(Debit) as Debit ,count(Debit) as PayCount from Supplier_Accounts where Transaction_Type='Payment'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') and Tender_Type = 'Cash'"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(1, 1) = Format(Val(rs.Fields("Debit") & "") * -1, "0.00")
        grdTrans.TextMatrix(1, 2) = rs.Fields("PayCount")
    End If
    rs.Close
    ActiveReadServer "SELECT Extra,Count(Extra) as Tend_Count" & _
    ",SUM(ISNULL(Sales_Journal.Sales_Tax, 0)) AS Sales_Tax" & _
    ",SUM(ISNULL(Sales_Journal.Line_Total, 0)) AS Line_Total" & _
    ",SUM(ISNULL(Sales_Journal.Ave_Cost, 0)) AS Ave_Cost" & _
    " From Sales_Journal" & _
    " WHERE (Sales_Journal.Function_Key IN (7)) and Extra in ('Corr','Void','Return Item','Wastage')" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " GROUP BY Extra"
    While Not rs.EOF
        Select Case rs.Fields("Extra") & ""
            Case ""
                grdStock.TextMatrix(2, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
            Case "Corr"
                grdCount.TextMatrix(1, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(1, 2) = rs.Fields("Tend_Count")
            Case "Void"
                grdCount.TextMatrix(2, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(2, 2) = rs.Fields("Tend_Count")
            Case "Return Item"
                grdCount.TextMatrix(3, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(3, 2) = rs.Fields("Tend_Count")
            Case "Wastage"
                grdCount.TextMatrix(4, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(4, 2) = rs.Fields("Tend_Count")
        End Select
        rs.MoveNext
    Wend
    rs.Close
    
    grdStock.TextMatrix(1, 1) = "0.00"
    grdStock.TextMatrix(2, 1) = "0.00"
    grdStock.TextMatrix(3, 1) = "0.00"
    grdStock.TextMatrix(4, 1) = "0.00"
    grdStock.TextMatrix(5, 1) = "0.00"
    grdStock.TextMatrix(6, 1) = "0.00"
    
    Stock_Value = 0
    ActiveReadServer "Select * from Stock_Value" & _
    " WHERE Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'"
    While Not rs.EOF
        Stock_Value = Stock_Value + rs.Fields("Stock_Value")
        rs.MoveNext
    Wend
    rs.Close
    
    Variance = 0
    ActiveReadServer "Select sum(Variance*Ave_Cost) as Variance" & _
    " from Stock_Take_Variance where " & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        Variance = Val(rs.Fields("Variance") & "")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Credit) as Deposit_Value,count(Credit) as Deposit_Count from Room_Accounts where Transaction_Type = 'Deposit' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(2, 1) = Format(Val(grdTrans.ValueMatrix(2, 1) + Val(rs.Fields("Deposit_Value") & "")), "0.00")
        grdTrans.TextMatrix(2, 2) = Val(grdTrans.ValueMatrix(2, 2) + rs.Fields("Deposit_Count") & "")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Credit) as Receipt_Value,count(Credit) as Receipt_Count from Room_Accounts where Transaction_Type ='Receipt' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(3, 1) = Format(Val(grdTrans.ValueMatrix(3, 1) + Val(rs.Fields("Receipt_Value") & "")), "0.00")
        grdTrans.TextMatrix(3, 2) = Val(grdTrans.ValueMatrix(3, 2) + rs.Fields("Receipt_Count") & "")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Credit) as Receipt_Value,count(Credit) as Receipt_Count from Debtor_Accounts where Transaction_Type ='Receipt' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(3, 1) = Format(Val(grdTrans.ValueMatrix(3, 1) + Val(rs.Fields("Receipt_Value") & "")), "0.00")
        grdTrans.TextMatrix(3, 2) = Val(grdTrans.ValueMatrix(3, 2) + rs.Fields("Receipt_Count") & "")
    End If
    rs.Close
    grdTrans.TextMatrix(4, 1) = Format(grdTrans.ValueMatrix(4, 1) + (grdTrans.ValueMatrix(1, 1) + grdTrans.ValueMatrix(2, 1) + grdTrans.ValueMatrix(3, 1)), "0.00")
    grdTrans.TextMatrix(4, 2) = grdTrans.ValueMatrix(4, 2) + (grdTrans.ValueMatrix(1, 2) + grdTrans.ValueMatrix(2, 2) + grdTrans.ValueMatrix(3, 2))
    grdTrans.TextMatrix(5, 2) = grdTrans.ValueMatrix(4, 2) + grdRev.ValueMatrix(6, 2)
    grdTrans.TextMatrix(5, 1) = Format(grdTrans.ValueMatrix(4, 1) + grdRev.ValueMatrix(6, 1), "0.00")
    grdStock.TextMatrix(2, 1) = "0.00"

    
   
       ActiveReadServer "SELECT SUM(Ave_Cost*Qty) AS Ave_Cost" & _
    " From Sales_Journal" & _
    " Where Function_Key =7  and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') "

    grdStock.TextMatrix(2, 1) = Variance * -1
    
    If rs.RecordCount > 0 Then
      grdStock.TextMatrix(2, 1) = Format(Val(rs.Fields("Ave_Cost") & "") + grdStock.ValueMatrix(2, 1), "0.00")
      Ave_Cost = Val(rs.Fields("Ave_Cost") & "")
    End If
    rs.Close
    
    Ave_Cost = Ave_Cost - Variance
    ActiveReadServer "Select sum(Qty*Ave_Cost) as Trans_Value from Transfer_Journal where" & _
    " Department_No like '" & DeptString & "'" & _
    " and Trans_Location_No like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdStock.TextMatrix(3, 1) = Format(Val(grdStock.ValueMatrix(3, 1) + rs.Fields("Trans_Value") & ""), "0.00")
    End If
    rs.Close
    
    grdCred.TextMatrix(1, 3) = "0.00"
    
    
    '**********************************************************
    'ABYC Journal from debtors
    'and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    grdCred.TextMatrix(2, 3) = "0.00"
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    'ABYC Journal from debtors
    grdCred.TextMatrix(3, 3) = "0.00"
    grdCred.TextMatrix(4, 3) = "0.00"
    grdCred.TextMatrix(5, 3) = "0.00"
    
    ActiveReadServer "Select sum(Line_Total) as Pur_Value,SUM((Line_Total*((100+Vat_Rate)/100)-Line_Total)) as Pur_Tax from Purchase_Journal where" & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'" & _
    " and Invoice_No is not Null and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdStock.TextMatrix(5, 1) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdCred.TextMatrix(2, 3) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdCred.TextMatrix(5, 3) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdTax.TextMatrix(5, 1) = Format(Val(rs.Fields("Pur_Tax") & ""), "0.00")
    End If
    rs.Close
    
    
    ActiveReadServer "Select sum(Debit-Credit) as Journal from Debtor_Accounts where Transaction_Type ='Journal' and (Date_Time > '" & mthViewStart.Value - 1 & "'  and Date_Time <'" & Selender & "')"
    'and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time <'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    'and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
    grdCred.TextMatrix(3, 1) = Format(Val(rs.Fields("Journal") & ""), "0.00")
    End If
    rs.Close
    
    
    
    
    ActiveReadServer "Select sum(Credit*-1) as Credit from Supplier_Accounts where Transaction_Type ='Supplier Invoice' and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(2, 3) = Format(Val(grdCred.TextMatrix(2, 3)) + Val(rs.Fields("Credit") & ""), "0.00")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Debit-Credit) as Journal from Supplier_Accounts where Transaction_Type ='Journal' and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        'ABYC Journal from debtors
        grdCred.TextMatrix(3, 3) = Format(Val(rs.Fields("Journal") & ""), "0.00")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Debit) as Debit from Supplier_Accounts where Transaction_Type ='Payment' and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(4, 3) = Format(Val(rs.Fields("Debit") & ""), "0.00")
    End If
    rs.Close
    
    ActiveReadServer "Select Sum(Balance) as Balance from Suppliers "
    If rs.RecordCount > 0 Then
    grdCred.TextMatrix(5, 3) = Format(Val(rs.Fields("Balance") & ""), "0.00")
    End If
    rs.Close
    
    grdCred.TextMatrix(1, 3) = Format(Round(grdCred.ValueMatrix(2, 3) + grdCred.ValueMatrix(3, 3) + grdCred.ValueMatrix(4, 3) - grdCred.ValueMatrix(5, 3), 2), "0.00")
    
    
    
    
    
    
    
    
    
    
    ActiveReadServer "Select sum(Balance) as Balance from Debtors"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(5, 1) = Format(Val(rs.Fields("Balance") & ""), "0.00")
    End If
    rs.Close
    grdCred.TextMatrix(2, 1) = Format(Val(grdRev.TextMatrix(4, 1) & ""), "0.00")
    grdCred.TextMatrix(4, 1) = Format(Val(grdTrans.TextMatrix(3, 1) & ""), "0.00")
    
    ActiveReadServer "Select sum(Qty*Ave_Cost) as Trans_Value from Transfer_Journal where" & _
    " Department_No like '" & DeptString & "'" & _
    " and Rec_Location_No like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdStock.TextMatrix(4, 1) = Format(Val(grdStock.ValueMatrix(4, 1) + rs.Fields("Trans_Value") & ""), "0.00")
    End If
    rs.Close
    
    grdStock.TextMatrix(6, 1) = Format(Stock_Value, "0.00")
    
    Open_Stock_Value = grdStock.ValueMatrix(6, 1) - grdStock.ValueMatrix(5, 1) - grdStock.ValueMatrix(4, 1) + grdStock.ValueMatrix(3, 1) + grdStock.ValueMatrix(2, 1)
    
    grdStock.TextMatrix(1, 1) = Format(Open_Stock_Value, "0.00")
    
    
    grdGP.TextMatrix(0, 1) = "0.00%"
    grdGP.TextMatrix(1, 1) = "0.00"
    grdGP.TextMatrix(0, 3) = "0"
    grdGP.TextMatrix(1, 3) = "0.00"
    If Ave_Cost < 0 Then grdGP.TextMatrix(1, 1) = Format((grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) + Ave_Cost, "0.00")
    If Ave_Cost > 0 Then grdGP.TextMatrix(1, 1) = Format((grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) - Ave_Cost, "0.00")
    
    ''grdGP.TextMatrix(1, 1) = Format((grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) - Ave_Cost, "0.00")
    
    If grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1) <> 0 Then
        grdGP.TextMatrix(0, 1) = Round((grdGP.ValueMatrix(1, 1) / (grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) * 100), 3) & "%"
    End If
    ActiveReadServer "Select sum(Covers) as Covers  from Sales_Journal where Function_Key IN ( 9, 10, 11, 12, 13) and " & _
    " ((Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "'))"
    If rs.RecordCount > 0 Then
        grdGP.TextMatrix(0, 3) = Int(Val(rs.Fields("Covers") & ""))
        If Int(rs.Fields("Covers")) <> 0 Then
            grdGP.TextMatrix(1, 3) = Format((grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) / Int(rs.Fields("Covers")), "0.00")
        End If
    End If
    rs.Close
    ActiveReadServer "Select sum(Tipp) as Tipp,Sum(Tipp_Count) as Tipp_Count,Sum(Discount_Perc_Value) as Discount,Sum(Discount_Perc_Qty) as Discount_Qty, Sum(Discount_amt_Value) as Discount_amt,Sum(Discount_amt_Qty) as Discount_amt_Qty from Counters " & _
        "where Cashup_No in " & _
        "(Select Cashup_no from sales_journal where Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCount.TextMatrix(7, 1) = Format(Val(rs.Fields("Tipp") & "") * factor, "0.00")
        grdCount.TextMatrix(5, 1) = Format(Val(rs.Fields("Discount") & "") * factor, "0.00")
        grdCount.TextMatrix(6, 1) = Format(Val(rs.Fields("Discount_amt") & "") * factor, "0.00")
        grdCount.TextMatrix(7, 2) = Val(rs.Fields("Tipp_Count") & "") * factor
        grdCount.TextMatrix(5, 2) = Val(rs.Fields("Discount_Qty") & "") * factor
        grdCount.TextMatrix(6, 2) = Val(rs.Fields("Discount_amt_Qty") & "") * factor
    End If
    rs.Close
End Sub


Private Sub Daytratealllocationsalldepartments()

LocString = "%"
DeptString = "%"
grdTrans.TextMatrix(1, 1) = "0.00"
    grdTrans.TextMatrix(1, 2) = "0"
    grdTrans.TextMatrix(2, 1) = "0.00"
    grdTrans.TextMatrix(2, 2) = "0"
    grdTrans.TextMatrix(3, 1) = "0.00"
    grdTrans.TextMatrix(3, 2) = "0"
    grdTrans.TextMatrix(4, 1) = "0.00"
    grdTrans.TextMatrix(4, 2) = "0"
    grdTrans.TextMatrix(5, 2) = "0"
    grdTrans.TextMatrix(5, 1) = "0.00"
    
    grdRev.TextMatrix(1, 1) = "0.00"
    grdRev.TextMatrix(2, 1) = "0.00"
    grdRev.TextMatrix(3, 1) = "0.00"
    grdRev.TextMatrix(4, 1) = "0.00"
    grdRev.TextMatrix(5, 1) = "0.00"
    grdRev.TextMatrix(6, 1) = "0.00"
    grdRev.TextMatrix(1, 2) = "0"
    grdRev.TextMatrix(2, 2) = "0"
    grdRev.TextMatrix(3, 2) = "0"
    grdRev.TextMatrix(4, 2) = "0"
    grdRev.TextMatrix(5, 2) = "0"
    grdRev.TextMatrix(6, 2) = "0"
    grdTax.TextMatrix(1, 1) = "0.00"
    grdTax.TextMatrix(2, 1) = "0.00"
    grdTax.TextMatrix(3, 1) = "0.00"
    grdTax.TextMatrix(4, 1) = "0.00"
    grdTax.TextMatrix(5, 1) = "0.00"
    grdStock.TextMatrix(2, 1) = "0.00"
 

        ActiveReadServer "SELECT Function_Key," & _
        "Case Function_Key" & _
            " WHEN '9' THEN 'Cash'" & _
            " WHEN '10' THEN 'Card'" & _
            " WHEN '11' THEN 'Voucher'" & _
            " WHEN '12' THEN 'Charge'" & _
            " WHEN '13' THEN 'Loyalty'" & _
        " END As Sale_Type" & _
        ",COUNT(Line_No) AS Tend_Count,SUM(Sales_Tax) AS Sales_Tax,SUM(Line_Total) AS Line_Total,Sum(Ave_Cost) as Ave_Cost" & _
        " From Sales_Journal" & _
        " WHERE (Function_Key IN ( 9, 10, 11, 12, 13)) " & _
        " and Department_No like '" & DeptString & "'" & _
        " and Location like '" & LocString & "'" & _
        " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
        " GROUP BY Function_Key"
        'x = rs.RecordCount
'      If rs.RecordCount > 1 Then rs.MoveFirst
      

While Not rs.EOF
        Select Case rs.Fields("Sale_Type")
            Case "Cash"
                grdRev.TextMatrix(1, 1) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdRev.TextMatrix(1, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Card"
                grdRev.TextMatrix(2, 1) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdRev.TextMatrix(2, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Voucher"
                grdRev.TextMatrix(3, 1) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdRev.TextMatrix(3, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Charge"
                grdRev.TextMatrix(4, 1) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdRev.TextMatrix(4, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Loyalty"
                grdRev.TextMatrix(5, 1) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdRev.TextMatrix(5, 2) = Val(rs.Fields("Tend_Count") & "")
        End Select
    grdRev.TextMatrix(6, 1) = Format(grdRev.ValueMatrix(6, 1) + Round(Val(rs.Fields("Line_Total")), 2), "0.00")
    grdRev.TextMatrix(6, 2) = grdRev.ValueMatrix(6, 2) + Val(rs.Fields("Tend_Count") & "")

    grdTax.TextMatrix(3, 1) = Format(grdTax.ValueMatrix(3, 1) + Round(Val(rs.Fields("Sales_Tax") & ""), 2), "0.00")
    grdTax.TextMatrix(4, 1) = Format(grdTax.ValueMatrix(4, 1) + Round(Val(rs.Fields("Sales_Tax") & ""), 2), "0.00")
    



rs.MoveNext
Wend
rs.Close


    ActiveReadServer "Select isnull(SUM(Line_Total),0) AS Taxable" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20)) and Sales_Tax <> 0" & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"

    grdTax.TextMatrix(1, 1) = Format(grdTax.ValueMatrix(1, 1) + Round(Val(rs.Fields("Taxable") & ""), 2), "0.00")
   
    rs.Close
    
    ActiveReadServer "Select isnull(SUM(Line_Total),0) AS Taxable" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20)) and Sales_Tax = 0" & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"

    grdTax.TextMatrix(2, 1) = Format(grdTax.ValueMatrix(2, 1) + Round(Val(rs.Fields("Taxable") & ""), 2), "0.00")
    
    
    rs.Close
    
    
    If Branch_Type = 5 And Selender > Date Then
        picLive.Visible = True
        ActiveReadServer "Select sum(isnull(line_Total,0)) as Total from Table_Listing where isnull(Extra_Function,'') = ''"
        If rs.Fields("Total") = 0 Then
            lblLive.Caption = "You have no Open Tables"
        Else
            lblLive.Caption = "Total with Open Tables: " & Format(rs.Fields("Total") + grdRev.ValueMatrix(6, 1), "0.00")
        End If
        rs.Close
    End If

    grdTrans.TextMatrix(5, 2) = "0"
    grdTrans.TextMatrix(5, 1) = "0.00"

    grdCount.TextMatrix(1, 1) = "0.00"
    grdCount.TextMatrix(2, 1) = "0.00"
    grdCount.TextMatrix(3, 1) = "0.00"
    grdCount.TextMatrix(4, 1) = "0.00"
    grdCount.TextMatrix(5, 1) = "0.00"
    grdCount.TextMatrix(6, 1) = "0.00"
    grdCount.TextMatrix(7, 1) = "0.00"
    grdCount.TextMatrix(1, 2) = "0"
    grdCount.TextMatrix(2, 2) = "0"
    grdCount.TextMatrix(3, 2) = "0"
    grdCount.TextMatrix(4, 2) = "0"
    grdCount.TextMatrix(5, 2) = "0"
    grdCount.TextMatrix(6, 2) = "0"
    grdCount.TextMatrix(7, 2) = "0"

    ActiveReadServer "Select sum(Debit) as Debit ,count(Debit) as PayCount from Supplier_Accounts where Transaction_Type='Payment'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') and Tender_Type = 'Cash'"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(1, 1) = Format(Val(rs.Fields("Debit") & "") * -1, "0.00")
        grdTrans.TextMatrix(1, 2) = rs.Fields("PayCount")
    End If
    rs.Close
    ActiveReadServer "SELECT Extra,Count(Extra) as Tend_Count" & _
    ",SUM(ISNULL(Sales_Journal.Sales_Tax, 0)) AS Sales_Tax" & _
    ",SUM(ISNULL(Sales_Journal.Line_Total, 0)) AS Line_Total" & _
    ",SUM(ISNULL(Sales_Journal.Ave_Cost, 0)) AS Ave_Cost" & _
    " From Sales_Journal" & _
    " WHERE (Sales_Journal.Function_Key IN (7)) and Extra in ('Corr','Void','Return Item','Wastage')" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " GROUP BY Extra"
    While Not rs.EOF
        Select Case rs.Fields("Extra") & ""
            Case ""
                grdStock.TextMatrix(2, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
            Case "Corr"
                grdCount.TextMatrix(1, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(1, 2) = rs.Fields("Tend_Count")
            Case "Void"
                grdCount.TextMatrix(2, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(2, 2) = rs.Fields("Tend_Count")
            Case "Return Item"
                grdCount.TextMatrix(3, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(3, 2) = rs.Fields("Tend_Count")
            Case "Wastage"
                grdCount.TextMatrix(4, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(4, 2) = rs.Fields("Tend_Count")
        End Select
        rs.MoveNext
    Wend
    rs.Close

    grdStock.TextMatrix(1, 1) = "0.00"
    grdStock.TextMatrix(2, 1) = "0.00"
    grdStock.TextMatrix(3, 1) = "0.00"
    grdStock.TextMatrix(4, 1) = "0.00"
    grdStock.TextMatrix(5, 1) = "0.00"
    grdStock.TextMatrix(6, 1) = "0.00"

    Stock_Value = 0
    ActiveReadServer "Select * from Stock_Value" & _
    " WHERE Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'"
    While Not rs.EOF
        Stock_Value = Stock_Value + rs.Fields("Stock_Value")
        rs.MoveNext
    Wend
    rs.Close

  ActiveReadServer "Select sum(Credit) as Deposit_Value,count(Credit) as Deposit_Count from Room_Accounts where Transaction_Type = 'Deposit' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(2, 1) = Format(Val(grdTrans.ValueMatrix(2, 1) + Val(rs.Fields("Deposit_Value") & "")), "0.00")
        grdTrans.TextMatrix(2, 2) = Val(grdTrans.ValueMatrix(2, 2) + rs.Fields("Deposit_Count") & "")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Credit) as Receipt_Value,count(Credit) as Receipt_Count from Room_Accounts where Transaction_Type ='Receipt' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(3, 1) = Format(Val(grdTrans.ValueMatrix(3, 1) + Val(rs.Fields("Receipt_Value") & "")), "0.00")
        grdTrans.TextMatrix(3, 2) = Val(grdTrans.ValueMatrix(3, 2) + rs.Fields("Receipt_Count") & "")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Credit) as Receipt_Value,count(Credit) as Receipt_Count from Debtor_Accounts where Transaction_Type ='Receipt' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(3, 1) = Format(Val(grdTrans.ValueMatrix(3, 1) + Val(rs.Fields("Receipt_Value") & "")), "0.00")
        grdTrans.TextMatrix(3, 2) = Val(grdTrans.ValueMatrix(3, 2) + rs.Fields("Receipt_Count") & "")
    End If
    rs.Close
    grdTrans.TextMatrix(4, 1) = Format(grdTrans.ValueMatrix(4, 1) + (grdTrans.ValueMatrix(1, 1) + grdTrans.ValueMatrix(2, 1) + grdTrans.ValueMatrix(3, 1)), "0.00")
    grdTrans.TextMatrix(4, 2) = grdTrans.ValueMatrix(4, 2) + (grdTrans.ValueMatrix(1, 2) + grdTrans.ValueMatrix(2, 2) + grdTrans.ValueMatrix(3, 2))
    grdTrans.TextMatrix(5, 2) = grdTrans.ValueMatrix(4, 2) + grdRev.ValueMatrix(6, 2)
    grdTrans.TextMatrix(5, 1) = Format(grdTrans.ValueMatrix(4, 1) + grdRev.ValueMatrix(6, 1), "0.00")
    grdStock.TextMatrix(2, 1) = "0.00"

    

       ActiveReadServer "SELECT SUM(Ave_Cost*Qty) AS Ave_Cost" & _
    " From Sales_Journal" & _
    " Where Function_Key =7 and   Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') "

     If rs.RecordCount > 0 Then
      grdStock.TextMatrix(2, 1) = Format(Val(rs.Fields("Ave_Cost") & "") + grdStock.ValueMatrix(2, 1), "0.00")
      Ave_Cost = Val(rs.Fields("Ave_Cost") & "")
    End If
    rs.Close
grdCred.TextMatrix(1, 3) = "0.00"
    grdCred.TextMatrix(2, 3) = "0.00"
    grdCred.TextMatrix(3, 3) = "0.00"
    'grdCred.TextMatrix(4, 3) = "0.00"
    
    grdCred.TextMatrix(5, 3) = "0.00"
     ActiveReadServer "Select sum(Line_Total) as Pur_Value,SUM((Line_Total*((100+Vat_Rate)/100)-Line_Total)) as Pur_Tax from Purchase_Journal where" & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'" & _
    " and Invoice_No is not Null and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdStock.TextMatrix(5, 1) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdCred.TextMatrix(2, 3) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdCred.TextMatrix(5, 3) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdTax.TextMatrix(5, 1) = Format(Val(rs.Fields("Pur_Tax") & ""), "0.00")
    End If
    rs.Close
    
    
     ActiveReadServer "Select sum(Debit-Credit) as Journal from Debtor_Accounts where Transaction_Type ='Journal' and (Date_Time > '" & mthViewStart.Value - 1 & "'  and Date_Time <'" & Selender & "')"
    'and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time <'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    'and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
    grdCred.TextMatrix(3, 1) = Format(Val(rs.Fields("Journal") & ""), "0.00")
    End If
    rs.Close
    
    
    
    ActiveReadServer "Select sum(Debit-Credit) as Journal from Supplier_Accounts where Transaction_Type ='Journal' and (Date_Time > '" & mthViewStart.Value - 1 & "'  and Date_Time <'" & Selender & "')"
    'and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time <'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    'and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
    grdCred.TextMatrix(3, 3) = Format(Val(rs.Fields("Journal") & ""), "0.00")
    End If
    rs.Close
    
    
    
    
    
    
    
    
    ActiveReadServer "Select sum(Credit*-1) as Credit from Supplier_Accounts where Transaction_Type ='Supplier Invoice' and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(2, 3) = Format(Val(rs.Fields("Credit") & ""), "0.00")
    End If
    rs.Close
    
'    ActiveReadServer "Select sum(Debit-Credit) as Journal from Supplier_Accounts where Transaction_Type ='Journal' and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
'    If rs.RecordCount > 0 Then
'        grdCred.TextMatrix(3, 3) = Format(Val(rs.Fields("Journal") & ""), "0.00")
'    End If
'    rs.Close
'
    
    
    'Payment suppliersaccounts
    ActiveReadServer "Select sum(Debit) as Debit from Supplier_Accounts where Transaction_Type ='Payment' and (Date_Time > '" & mthViewStart.Value - 1 & "'  and Date_Time <'" & Selender & "')"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(4, 3) = Format(Val(rs.Fields("Debit") & ""), "0.00")
    End If
    rs.Close
    
    
    
    
    
    
    
    
    
    ActiveReadServer "Select Sum(Balance) as Balance from Suppliers "
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(5, 3) = Format(Val(rs.Fields("Balance") & ""), "0.00")
    End If
    rs.Close
    
    grdCred.TextMatrix(1, 3) = Format(Round(grdCred.ValueMatrix(2, 3) + grdCred.ValueMatrix(3, 3) + grdCred.ValueMatrix(4, 3) - grdCred.ValueMatrix(5, 3), 2), "0.00")
    
    ActiveReadServer "Select sum(Balance) as Balance from Debtors"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(5, 1) = Format(Val(rs.Fields("Balance") & ""), "0.00")
    End If
    rs.Close
    grdCred.TextMatrix(2, 1) = Format(Val(grdRev.TextMatrix(4, 1) & ""), "0.00")
    grdCred.TextMatrix(4, 1) = Format(Val(grdTrans.TextMatrix(3, 1) & ""), "0.00")
    
    ActiveReadServer "Select sum(Qty*Ave_Cost) as Trans_Value from Transfer_Journal where" & _
    " Department_No like '" & DeptString & "'" & _
    " and Rec_Location_No like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdStock.TextMatrix(4, 1) = Format(Val(grdStock.ValueMatrix(4, 1) + rs.Fields("Trans_Value") & ""), "0.00")
    End If
    rs.Close
    
    grdStock.TextMatrix(6, 1) = Format(Stock_Value, "0.00")
    
    Open_Stock_Value = grdStock.ValueMatrix(6, 1) - grdStock.ValueMatrix(5, 1) - grdStock.ValueMatrix(4, 1) + grdStock.ValueMatrix(3, 1) + grdStock.ValueMatrix(2, 1)
    
    grdStock.TextMatrix(1, 1) = Format(Open_Stock_Value, "0.00")
    
    
    grdGP.TextMatrix(0, 1) = "0.00%"
    grdGP.TextMatrix(1, 1) = "0.00"
    grdGP.TextMatrix(0, 3) = "0"
    grdGP.TextMatrix(1, 3) = "0.00"
    
    grdGP.TextMatrix(1, 1) = Format((grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) - Ave_Cost, "0.00")
    
    If grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1) <> 0 Then
        grdGP.TextMatrix(0, 1) = Round((grdGP.ValueMatrix(1, 1) / (grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) * 100), 3) & "%"
    End If
    ActiveReadServer "Select sum(Covers) as Covers  from Sales_Journal where Function_Key IN ( 9, 10, 11, 12, 13) and " & _
    " ((Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "'))"
    If rs.RecordCount > 0 Then
        grdGP.TextMatrix(0, 3) = Int(Val(rs.Fields("Covers") & ""))
        If Int(rs.Fields("Covers")) <> 0 Then
            grdGP.TextMatrix(1, 3) = Format((grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) / Int(rs.Fields("Covers")), "0.00")
        End If
    End If
    rs.Close
    If factor = 0 Then factor = 1
    
    '***************************************
    'Kotie 15-03-2013 06:39
    'Fix Discount amount not on report
    '***************************************
    ActiveReadServer "Select sum(Tipp) as Tipp,Sum(Tipp_Count) as Tipp_Count,Sum(Discount_Perc_Value) as Discount,Sum(Discount_Perc_Qty) as Discount_Qty, Sum(Discount_amt_Value) as Discount_amt,Sum(Discount_amt_Qty) as Discount_amt_Qty from Counters where Cashup_No in (Select Cashup_no from sales_journal where Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCount.TextMatrix(7, 1) = Format(Val(rs.Fields("Tipp") & "") * factor, "0.00")
        grdCount.TextMatrix(5, 1) = Format(Val(rs.Fields("Discount") & "") * factor, "0.00")
        grdCount.TextMatrix(6, 1) = Format(Val(rs.Fields("Discount_amt") & "") * factor, "0.00")
        grdCount.TextMatrix(7, 2) = Format(Val(rs.Fields("Tipp_Count") & "") * factor, "0")
        grdCount.TextMatrix(5, 2) = Val(rs.Fields("Discount_Qty") & "") * factor
        grdCount.TextMatrix(6, 2) = Val(rs.Fields("Discount_amt_Qty") & "") * factor
    End If
    rs.Close
    
'    For ff = 1 To 10
'    grdCred.TextMatrix(3, ff) = ff
'    Next ff

End Sub




Private Sub Day_trade()
If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    
   If cmb1.Text = "<All Locations>" And cmb3.Text = "<All Departments>" Then
   Daytratealllocationsalldepartments
   Exit Sub
   End If
    
    
    
    
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdTrans.TextMatrix(1, 1) = "0.00"
    grdTrans.TextMatrix(1, 2) = "0"
    grdTrans.TextMatrix(2, 1) = "0.00"
    grdTrans.TextMatrix(2, 2) = "0"
    grdTrans.TextMatrix(3, 1) = "0.00"
    grdTrans.TextMatrix(3, 2) = "0"
    grdTrans.TextMatrix(4, 1) = "0.00"
    grdTrans.TextMatrix(4, 2) = "0"
    grdTrans.TextMatrix(5, 2) = "0"
    grdTrans.TextMatrix(5, 1) = "0.00"
    
    grdRev.TextMatrix(1, 1) = "0.00"
    grdRev.TextMatrix(2, 1) = "0.00"
    grdRev.TextMatrix(3, 1) = "0.00"
    grdRev.TextMatrix(4, 1) = "0.00"
    grdRev.TextMatrix(5, 1) = "0.00"
    grdRev.TextMatrix(6, 1) = "0.00"
    grdRev.TextMatrix(1, 2) = "0"
    grdRev.TextMatrix(2, 2) = "0"
    grdRev.TextMatrix(3, 2) = "0"
    grdRev.TextMatrix(4, 2) = "0"
    grdRev.TextMatrix(5, 2) = "0"
    grdRev.TextMatrix(6, 2) = "0"
    grdTax.TextMatrix(1, 1) = "0.00"
    grdTax.TextMatrix(2, 1) = "0.00"
    grdTax.TextMatrix(3, 1) = "0.00"
    grdTax.TextMatrix(4, 1) = "0.00"
    grdTax.TextMatrix(5, 1) = "0.00"
    grdStock.TextMatrix(2, 1) = "0.00"
    factor = 1
    
   'To Use in comparison
     ActiveReadServer "Select isnull(" & _
    " (SELECT SUM(Line_Total) AS Line_Total" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20))" & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " )/" & _
    " (SELECT SUM(Line_Total) AS Line_Total" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20))" & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '%'" & _
    " and Location like '%'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " ),0) as Factor"
    
    
  
    
    factor = rs.Fields("Factor")
    rs.Close
    ActiveReadServer "SELECT Function_Key," & _
        "Case Function_Key" & _
            " WHEN 9 THEN 'Cash'" & _
            " WHEN 10 THEN 'Card'" & _
            " WHEN 11 THEN 'Voucher'" & _
            " WHEN 12 THEN 'Charge'" & _
            " WHEN 13 THEN 'Loyalty'" & _
        " END As Sale_Type" & _
        ",COUNT(Line_No) AS Tend_Count,SUM(Sales_Tax) AS Sales_Tax,SUM(Line_Total) AS Line_Total,Sum(Ave_Cost) as Ave_Cost" & _
        " From Sales_Journal" & _
        " WHERE (Function_Key IN ( 9, 10, 11, 12, 13)) " & _
        " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
        " GROUP BY Function_Key"
    While Not rs.EOF
        Select Case rs.Fields("Sale_Type")
            Case "Cash"
                grdRev.TextMatrix(1, 1) = "N/A" 'Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(1, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Card"
                grdRev.TextMatrix(2, 1) = "N/A" ' Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(2, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Voucher"
                grdRev.TextMatrix(3, 1) = "N/A" ' Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(3, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Charge"
                grdRev.TextMatrix(4, 1) = "N/A" ' Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(4, 2) = Val(rs.Fields("Tend_Count") & "")
            Case "Loyalty"
                grdRev.TextMatrix(5, 1) = "N/A" ' Format(Val(rs.Fields("Line_Total") * factor & ""), "0.00")
                grdRev.TextMatrix(5, 2) = Val(rs.Fields("Tend_Count") & "")
        End Select
        
        grdRev.TextMatrix(1, 1) = "N/A"
        grdRev.TextMatrix(2, 1) = "N/A"
        grdRev.TextMatrix(3, 1) = "N/A"
        grdRev.TextMatrix(4, 1) = "N/A"
        grdRev.TextMatrix(5, 1) = "N/A"
        
        
        
        
        grdRev.TextMatrix(6, 1) = Format(grdRev.ValueMatrix(6, 1) + Round(Val(rs.Fields("Line_Total") * factor & ""), 2), "0.00")
        grdRev.TextMatrix(6, 2) = grdRev.ValueMatrix(6, 2) + Val(rs.Fields("Tend_Count") & "")
                        
        grdTax.TextMatrix(3, 1) = Format(grdTax.ValueMatrix(3, 1) + Round(Val(rs.Fields("Sales_Tax") & "") * factor, 2), "0.00")
        grdTax.TextMatrix(4, 1) = Format(grdTax.ValueMatrix(4, 1) + Round(Val(rs.Fields("Sales_Tax") & "") * factor, 2), "0.00")
        
        rs.MoveNext
    Wend
    rs.Close
    
'    ActiveReadServer "Select isnull(SUM(Line_Total),0) AS Taxable" & _
'    " From Sales_Journal" & _
'    " Where (Function_Key in(7,20)) and Sales_Tax <> 0" & _
'    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
'    " and Department_No like '" & DeptString & "'" & _
'    " and Location like '" & LocString & "'" & _
'    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"

    
     ActiveReadServer "Select isnull(SUM(Line_Total),0) AS Taxable" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20)) and Sales_Tax <> 0 and Extra <> 'Corr' and Extra <> 'Void' " & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"

    
    
    
    
    
    
    
    
    
    
    
    
    
    grdTax.TextMatrix(1, 1) = Format(grdTax.ValueMatrix(1, 1) + Round(Val(rs.Fields("Taxable") & ""), 2), "0.00")
    rs.Close
    
    ActiveReadServer "Select isnull(SUM(Line_Total),0) AS Taxable" & _
    " From Sales_Journal" & _
    " Where (Function_Key in(7,20)) and Sales_Tax = 0" & _
    " and (isnull(Extra,'')='' or Extra in ('Return Item') or Function_Key =20)" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"

    grdTax.TextMatrix(2, 1) = Format(grdTax.ValueMatrix(2, 1) + Round(Val(rs.Fields("Taxable") & ""), 2), "0.00")
    rs.Close
    
    If Branch_Type = 5 And Selender > Date Then
        picLive.Visible = True
        ActiveReadServer "Select sum(isnull(line_Total,0)) as Total from Table_Listing where isnull(Extra_Function,'') = ''"
        If rs.Fields("Total") = 0 Then
            lblLive.Caption = "You have no Open Tables"
        Else
            lblLive.Caption = "Total with Open Tables: " & Format(rs.Fields("Total") + grdRev.ValueMatrix(6, 1), "0.00")
        End If
        rs.Close
    End If
    
    grdTrans.TextMatrix(5, 2) = "0"
    grdTrans.TextMatrix(5, 1) = "0.00"
    
    grdCount.TextMatrix(1, 1) = "0.00"
    grdCount.TextMatrix(2, 1) = "0.00"
    grdCount.TextMatrix(3, 1) = "0.00"
    grdCount.TextMatrix(4, 1) = "0.00"
    grdCount.TextMatrix(5, 1) = "0.00"
    grdCount.TextMatrix(6, 1) = "0.00"
    grdCount.TextMatrix(7, 1) = "0.00"
    grdCount.TextMatrix(1, 2) = "0"
    grdCount.TextMatrix(2, 2) = "0"
    grdCount.TextMatrix(3, 2) = "0"
    grdCount.TextMatrix(4, 2) = "0"
    grdCount.TextMatrix(5, 2) = "0"
    grdCount.TextMatrix(6, 2) = "0"
    grdCount.TextMatrix(7, 2) = "0"
    
    ActiveReadServer "Select sum(Debit) as Debit ,count(Debit) as PayCount from Supplier_Accounts where Transaction_Type='Payment'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') and Tender_Type = 'Cash'"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(1, 1) = Format(Val(rs.Fields("Debit") & "") * -1, "0.00")
        grdTrans.TextMatrix(1, 2) = rs.Fields("PayCount")
    End If
    rs.Close
    ActiveReadServer "SELECT Extra,Count(Extra) as Tend_Count" & _
    ",SUM(ISNULL(Sales_Journal.Sales_Tax, 0)) AS Sales_Tax" & _
    ",SUM(ISNULL(Sales_Journal.Line_Total, 0)) AS Line_Total" & _
    ",SUM(ISNULL(Sales_Journal.Ave_Cost, 0)) AS Ave_Cost" & _
    " From Sales_Journal" & _
    " WHERE (Sales_Journal.Function_Key IN (7)) and Extra in ('Corr','Void','Return Item','Wastage')" & _
    " and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
    " GROUP BY Extra"
    While Not rs.EOF
        Select Case rs.Fields("Extra") & ""
            Case ""
                grdStock.TextMatrix(2, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
            Case "Corr"
                grdCount.TextMatrix(1, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(1, 2) = rs.Fields("Tend_Count")
            Case "Void"
                grdCount.TextMatrix(2, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(2, 2) = rs.Fields("Tend_Count")
            Case "Return Item"
                grdCount.TextMatrix(3, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(3, 2) = rs.Fields("Tend_Count")
            Case "Wastage"
                grdCount.TextMatrix(4, 1) = Format(Abs(rs.Fields("Line_Total")), "0.00")
                grdCount.TextMatrix(4, 2) = rs.Fields("Tend_Count")
        End Select
        rs.MoveNext
    Wend
    rs.Close
    
    grdStock.TextMatrix(1, 1) = "0.00"
    grdStock.TextMatrix(2, 1) = "0.00"
    grdStock.TextMatrix(3, 1) = "0.00"
    grdStock.TextMatrix(4, 1) = "0.00"
    grdStock.TextMatrix(5, 1) = "0.00"
    grdStock.TextMatrix(6, 1) = "0.00"
    
    Stock_Value = 0
    ActiveReadServer "Select * from Stock_Value" & _
    " WHERE Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'"
    While Not rs.EOF
        Stock_Value = Stock_Value + rs.Fields("Stock_Value")
        rs.MoveNext
    Wend
    rs.Close
  
    
    ActiveReadServer "Select sum(Credit) as Deposit_Value,count(Credit) as Deposit_Count from Room_Accounts where Transaction_Type = 'Deposit' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(2, 1) = Format(Val(grdTrans.ValueMatrix(2, 1) + Val(rs.Fields("Deposit_Value") & "")), "0.00")
        grdTrans.TextMatrix(2, 2) = Val(grdTrans.ValueMatrix(2, 2) + rs.Fields("Deposit_Count") & "")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Credit) as Receipt_Value,count(Credit) as Receipt_Count from Room_Accounts where Transaction_Type ='Receipt' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(3, 1) = Format(Val(grdTrans.ValueMatrix(3, 1) + Val(rs.Fields("Receipt_Value") & "")), "0.00")
        grdTrans.TextMatrix(3, 2) = Val(grdTrans.ValueMatrix(3, 2) + rs.Fields("Receipt_Count") & "")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Credit) as Receipt_Value,count(Credit) as Receipt_Count from Debtor_Accounts where Transaction_Type ='Receipt' and " & _
    " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdTrans.TextMatrix(3, 1) = Format(Val(grdTrans.ValueMatrix(3, 1) + Val(rs.Fields("Receipt_Value") & "")), "0.00")
        grdTrans.TextMatrix(3, 2) = Val(grdTrans.ValueMatrix(3, 2) + rs.Fields("Receipt_Count") & "")
    End If
    rs.Close
    grdTrans.TextMatrix(4, 1) = Format(grdTrans.ValueMatrix(4, 1) + (grdTrans.ValueMatrix(1, 1) + grdTrans.ValueMatrix(2, 1) + grdTrans.ValueMatrix(3, 1)), "0.00")
    grdTrans.TextMatrix(4, 2) = grdTrans.ValueMatrix(4, 2) + (grdTrans.ValueMatrix(1, 2) + grdTrans.ValueMatrix(2, 2) + grdTrans.ValueMatrix(3, 2))
    grdTrans.TextMatrix(5, 2) = grdTrans.ValueMatrix(4, 2) + grdRev.ValueMatrix(6, 2)
    grdTrans.TextMatrix(5, 1) = Format(grdTrans.ValueMatrix(4, 1) + grdRev.ValueMatrix(6, 1), "0.00")
    grdStock.TextMatrix(2, 1) = "0.00"

  
     ActiveReadServer "SELECT SUM(Ave_Cost*Qty) AS Ave_Cost" & _
    " From Sales_Journal" & _
    " Where Function_Key =7  and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') "

  
  
    
    If rs.RecordCount > 0 Then
      grdStock.TextMatrix(2, 1) = Format(Val(rs.Fields("Ave_Cost") & "") + grdStock.ValueMatrix(2, 1), "0.00")
      Ave_Cost = Val(rs.Fields("Ave_Cost") & "")
    End If
    rs.Close

    
    grdCred.TextMatrix(1, 3) = "0.00"
    grdCred.TextMatrix(2, 3) = "0.00"
    grdCred.TextMatrix(3, 3) = "0.00"
    grdCred.TextMatrix(4, 3) = "0.00"
    grdCred.TextMatrix(5, 3) = "0.00"
    
    ActiveReadServer "Select sum(Line_Total) as Pur_Value,SUM((Line_Total*((100+Vat_Rate)/100)-Line_Total)) as Pur_Tax from Purchase_Journal where" & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'" & _
    " and Invoice_No is not Null and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdStock.TextMatrix(5, 1) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdCred.TextMatrix(2, 3) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdCred.TextMatrix(5, 3) = Format(Val(rs.Fields("Pur_Value") & ""), "0.00")
        grdTax.TextMatrix(5, 1) = Format(Val(rs.Fields("Pur_Tax") & ""), "0.00")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Credit*-1) as Credit from Supplier_Accounts where Transaction_Type ='Supplier Invoice' and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(2, 3) = Format(Val(rs.Fields("Credit") & ""), "0.00")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Debit-Credit) as Journal from Supplier_Accounts where Transaction_Type ='Journal' and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(3, 3) = Format(Val(rs.Fields("Journal") & ""), "0.00")
    End If
    rs.Close
    
    ActiveReadServer "Select sum(Debit) as Debit from Supplier_Accounts where Transaction_Type ='Payment' and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(4, 3) = Format(Val(rs.Fields("Debit") & ""), "0.00")
    End If
    rs.Close
    
    ActiveReadServer "Select Sum(Balance) as Balance from Suppliers "
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(5, 3) = Format(Val(rs.Fields("Balance") & ""), "0.00")
    End If
    rs.Close
    
    grdCred.TextMatrix(1, 3) = Format(Round(grdCred.ValueMatrix(2, 3) + grdCred.ValueMatrix(3, 3) + grdCred.ValueMatrix(4, 3) - grdCred.ValueMatrix(5, 3), 2), "0.00")
    
    ActiveReadServer "Select sum(Balance) as Balance from Debtors"
    If rs.RecordCount > 0 Then
        grdCred.TextMatrix(5, 1) = Format(Val(rs.Fields("Balance") & ""), "0.00")
    End If
    rs.Close
    grdCred.TextMatrix(2, 1) = Format(Val(grdRev.TextMatrix(4, 1) & ""), "0.00")
    grdCred.TextMatrix(4, 1) = Format(Val(grdTrans.TextMatrix(3, 1) & ""), "0.00")
    
    ActiveReadServer "Select sum(Qty*Ave_Cost) as Trans_Value from Transfer_Journal where" & _
    " Department_No like '" & DeptString & "'" & _
    " and Rec_Location_No like '" & LocString & "'" & _
    " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdStock.TextMatrix(4, 1) = Format(Val(grdStock.ValueMatrix(4, 1) + rs.Fields("Trans_Value") & ""), "0.00")
    End If
    rs.Close
    
    grdStock.TextMatrix(6, 1) = Format(Stock_Value, "0.00")
    
    Open_Stock_Value = grdStock.ValueMatrix(6, 1) - grdStock.ValueMatrix(5, 1) - grdStock.ValueMatrix(4, 1) + grdStock.ValueMatrix(3, 1) + grdStock.ValueMatrix(2, 1)
    
    grdStock.TextMatrix(1, 1) = Format(Open_Stock_Value, "0.00")
    
    
    grdGP.TextMatrix(0, 1) = "0.00%"
    grdGP.TextMatrix(1, 1) = "0.00"
    grdGP.TextMatrix(0, 3) = "0"
    grdGP.TextMatrix(1, 3) = "0.00"
    
    grdGP.TextMatrix(1, 1) = Format((grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) - Ave_Cost, "0.00")
    
    If grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1) <> 0 Then
        grdGP.TextMatrix(0, 1) = Round((grdGP.ValueMatrix(1, 1) / (grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) * 100), 3) & "%"
    End If
    ActiveReadServer "Select sum(Covers) as Covers  from Sales_Journal where Function_Key IN ( 9, 10, 11, 12, 13) and " & _
    " ((Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "'))"
    If rs.RecordCount > 0 Then
        grdGP.TextMatrix(0, 3) = Int(Val(rs.Fields("Covers") & ""))
        If Int(rs.Fields("Covers")) <> 0 Then
            grdGP.TextMatrix(1, 3) = Format((grdRev.ValueMatrix(6, 1) - grdTax.TextMatrix(3, 1)) / Int(rs.Fields("Covers")), "0.00")
        End If
    End If
    rs.Close
    
    
    
    If factor = 0 Then factor = 1
    '***************************************
    'Kotie 15-03-2013 06:39
    'Fix Discount amount not on report
    '***************************************
    ActiveReadServer "Select sum(Tipp) as Tipp,Sum(Tipp_Count) as Tipp_Count,Sum(Discount_Perc_Value) as Discount,Sum(Discount_Perc_Qty) as Discount_Qty, Sum(Discount_amt_Value) as Discount_amt,Sum(Discount_amt_Qty) as Discount_amt_Qty  from Counters where Cashup_No in (Select Cashup_no from sales_journal where Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    If rs.RecordCount > 0 Then
        grdCount.TextMatrix(7, 1) = Format(Val(rs.Fields("Tipp") & "") * factor, "0.00")
        grdCount.TextMatrix(5, 1) = Format(Val(rs.Fields("Discount") & "") * factor, "0.00")
        grdCount.TextMatrix(6, 1) = Format(Val(rs.Fields("Discount_amt") & "") * factor, "0.00")
        grdCount.TextMatrix(5, 2) = Val(rs.Fields("Discount_Qty") & "") * factor
        grdCount.TextMatrix(6, 2) = Val(rs.Fields("Discount_amt_Qty") & "") * factor
    If factor < 1 Then factor = Format(factor, "0.00")
    grdCount.TextMatrix(7, 2) = Format(Val(rs.Fields("Tipp_Count") & "") * factor, "0")
    
'    ActiveReadServer "Select sum(Tipp) as Tipp,Sum(Tipp_Count) as Tipp_Count,Sum(Discount_Perc_Value) as Discount,Sum(Discount_Perc_Qty) as Discount_Qty from Counters where Cashup_No in (Select Cashup_no from sales_journal where Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
'    If rs.RecordCount > 0 Then
'        grdCount.TextMatrix(7, 1) = Format(Val(rs.Fields("Tipp") & "") * factor, "0.00")
'        grdCount.TextMatrix(6, 1) = Format(Val(rs.Fields("Discount") & "") * factor, "0.00")
'
'        grdCount.TextMatrix(6, 2) = Val(rs.Fields("Discount_Qty") & "") * factor
'
'    If factor < 1 Then factor = Format(factor, "0.00")
'    grdCount.TextMatrix(7, 2) = Format(Val(rs.Fields("Tipp_Count") & "") * factor, "0")
    
    '***************************************
    
'    ActiveReadServer "Select sum(Tipp) as Tipp,Sum(Tipp_Count) as Tipp_Count,Sum(Discount_Perc_Value) as Discount,Sum(Discount_Perc_Qty) as Discount_Qty from Counters where Cashup_No in (Select Cashup_no from sales_journal where Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
'    If rs.RecordCount > 0 Then
'        grdCount.TextMatrix(7, 1) = Format(Val(rs.Fields("Tipp") & "") * factor, "0.00")
'        grdCount.TextMatrix(6, 1) = Format(Val(rs.Fields("Discount") & "") * factor, "0.00")
'        grdCount.TextMatrix(7, 2) = Val(rs.Fields("Tipp_Count") & "") * factor
'        grdCount.TextMatrix(6, 2) = Val(rs.Fields("Discount_Qty") & "") * factor
    End If
    rs.Close
    
End Sub


Private Sub Stock_Movement_Values()
    On Error GoTo trap
    Screen.MousePointer = 13
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, mthViewEnd.Value)
        lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Else
        Selender = mthViewEnd.Value
    End If
    If cmb1.Text <> "<All Locations>" Then
        LocString = Mid(cmb1.Text, 1, InStr(cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    grdTotal.TextMatrix(0, 2) = "0.00"
    grdTotal.TextMatrix(0, 3) = "0.00"
    grdTotal.TextMatrix(0, 4) = "0.00"
    grdTotal.TextMatrix(0, 5) = "0.00"
    grdTotal.TextMatrix(0, 6) = "0.00"
    grdTotal.TextMatrix(0, 7) = "0.00"
    grdTotal.TextMatrix(0, 8) = "0.00"
    grdTotal.TextMatrix(0, 9) = "0.00"
    grdTotal.TextMatrix(0, 10) = "0.00"
    grdTotal.TextMatrix(0, 11) = "0.00"
    grdTotal.TextMatrix(0, 12) = "0.00"
    grdMain.Rows = 1
    grdMain.SetFocus
    DoEvents
    ActiveReadServer3 "SELECT Products.Product_Code,Description From Products LEFT OUTER JOIN Quantities ON Products.Product_Code = Quantities.Product_Code where Location_No like '" & LocString & "' and Department_No like '" & DeptString & "' and Stock_Item=1 GROUP BY Products.Product_Code,Description order by Description"
    While Not rs3.EOF
         grdMain.Rows = grdMain.Rows + 1
         For i = 2 To grdMain.Cols - 1
            grdMain.TextMatrix(grdMain.Rows - 1, i) = "0"
        Next i
        ActiveReadServer1 "Exec Stock_Move_Values '" & DeptString & "', '" & LocString & "', '" & mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "', '" & mthViewEnd.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "', '" & rs3.Fields("Product_code") & "'"
            grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs1.Fields("Product_code")
            grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs1.Fields("Description")
            grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(Val(rs1.Fields("Qty_Received") & ""), "0.00")
            grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(Val(rs1.Fields("Qty_Trans_Out") & ""), "0.00")
            grdMain.TextMatrix(grdMain.Rows - 1, 7) = Format(Val(rs1.Fields("Qty_Trans_In") & ""), "0.00")
            grdMain.TextMatrix(grdMain.Rows - 1, 9) = Format(Val(rs1.Fields("Qty_Consumed") & ""), "0.00")
            grdMain.TextMatrix(grdMain.Rows - 1, 11) = Format(Val(rs1.Fields("Variance") & ""), "0.00")
            grdMain.TextMatrix(grdMain.Rows - 1, 12) = Format(Val(rs1.Fields("Stock_on_Hand") & ""), "0.00")
        rs1.Close
        grdMain.TextMatrix(grdMain.Rows - 1, 10) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 12) - grdMain.ValueMatrix(grdMain.Rows - 1, 11), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 10) + grdMain.ValueMatrix(grdMain.Rows - 1, 9) + grdMain.ValueMatrix(grdMain.Rows - 1, 6) - grdMain.ValueMatrix(grdMain.Rows - 1, 7) - grdMain.ValueMatrix(grdMain.Rows - 1, 3), "0.00")
        grdTotal.TextMatrix(0, 2) = Round(grdTotal.ValueMatrix(0, 2) + grdMain.ValueMatrix(grdMain.Rows - 1, 2), 3)
        grdTotal.TextMatrix(0, 3) = Round(grdTotal.ValueMatrix(0, 3) + grdMain.ValueMatrix(grdMain.Rows - 1, 3), 3)
        grdTotal.TextMatrix(0, 4) = Round(grdTotal.ValueMatrix(0, 4) + grdMain.ValueMatrix(grdMain.Rows - 1, 4), 3)
        grdTotal.TextMatrix(0, 5) = Round(grdTotal.ValueMatrix(0, 5) + grdMain.ValueMatrix(grdMain.Rows - 1, 5), 3)
        grdTotal.TextMatrix(0, 6) = Round(grdTotal.ValueMatrix(0, 6) + grdMain.ValueMatrix(grdMain.Rows - 1, 6), 3)
        grdTotal.TextMatrix(0, 7) = Round(grdTotal.ValueMatrix(0, 7) + grdMain.ValueMatrix(grdMain.Rows - 1, 7), 3)
        grdTotal.TextMatrix(0, 8) = Round(grdTotal.ValueMatrix(0, 8) + grdMain.ValueMatrix(grdMain.Rows - 1, 8), 3)
        grdTotal.TextMatrix(0, 9) = Round(grdTotal.ValueMatrix(0, 9) + grdMain.ValueMatrix(grdMain.Rows - 1, 9), 3)
        grdTotal.TextMatrix(0, 10) = Round(grdTotal.ValueMatrix(0, 10) + grdMain.ValueMatrix(grdMain.Rows - 1, 10), 3)
        grdTotal.TextMatrix(0, 11) = Round(grdTotal.ValueMatrix(0, 11) + grdMain.ValueMatrix(grdMain.Rows - 1, 11), 3)
        grdTotal.TextMatrix(0, 12) = Round(grdTotal.ValueMatrix(0, 12) + grdMain.ValueMatrix(grdMain.Rows - 1, 12), 3)
        DoEvents
        rs3.MoveNext
    Wend
    grdTotal.TextMatrix(0, 1) = " Products = " & rs3.RecordCount
    rs3.Close
    Screen.MousePointer = 0
    On Error GoTo 0
    Exit Sub
trap:
    rs3.Close
    rs1.Close
    Screen.MousePointer = 0
    On Error GoTo 0
End Sub
Private Sub cmb1_DropButtonClick()
    If cmb3.Visible = False Then
        cmb1.ListWidth = 0
    Else
        cmb1.ListWidth = 140
    End If
End Sub

Private Sub cmb1_GotFocus()
    If picDate.Visible = True Then Selection_Change
    picDate.Visible = False
    ButtonEx1.Value = Up
End Sub
Private Sub Transfer_Journal()
    picMain.Visible = True
    cmb1.Width = 2245
    cmb2.Width = 1800
    cmb3.Width = 2100
    cmb1.Left = 9240
    cmb2.Left = 7380
    cmb3.Left = 11530
    cmb3.Visible = True
    grdMain.Cols = 7
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Date & Time"
    grdMain.TextMatrix(0, 1) = "User"
    grdMain.TextMatrix(0, 2) = "Transfer No."
    grdMain.TextMatrix(0, 3) = "Transaction Type"
    grdMain.TextMatrix(0, 4) = "Transfering Location"
    grdMain.TextMatrix(0, 5) = "Receiving Location"
    grdMain.TextMatrix(0, 6) = "Line Total"
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignLeftCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignLeftCenter
    grdMain.ColAlignment(5) = flexAlignLeftCenter
    grdMain.ColAlignment(6) = flexAlignRightCenter
    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.15
    grdMain.ColWidth(2) = grdMain.Width * 0.075
    grdMain.ColWidth(3) = grdMain.Width * 0.15
    grdMain.ColWidth(4) = grdMain.Width * 0.19
    grdMain.ColWidth(5) = grdMain.Width * 0.19
    grdMain.ColWidth(6) = grdMain.Width * 0.095
    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    cmb1.AddItem "<Transfering Locations>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<Transfering Locations>"
    cmb3.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    cmb3.AddItem "<Receiving Locations>"
    While Not rs.EOF
        cmb3.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb3.Text = "<Receiving Locations>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Pre_Sales_Analysis()
    If cmb1.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    grdTotal.TextMatrix(0, 0) = ""
    grdTotal.TextMatrix(0, 1) = ""
    grdTotal.TextMatrix(0, 2) = ""
    grdTotal.TextMatrix(0, 3) = ""
    grdTotal.TextMatrix(0, 4) = ""
    grdTotal.TextMatrix(0, 5) = ""
    grdTotal.TextMatrix(0, 6) = ""
    ActiveReadServer "Select * from Price_Margin_View where " & _
    " Department_No like '" & DeptString & "' order by Product_Code"
    grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 1) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 2) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 3) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 4) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 5) = " Products = " & rs.RecordCount
    grdTotal.TextMatrix(0, 6) = " Products = " & rs.RecordCount
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Landed_Cost") & "", "0.00")
        If Val(rs.Fields("Selling_Price") & "") = 0 Then
            grdMain.TextMatrix(grdMain.Rows - 1, 3) = "0"
        Else
            grdMain.TextMatrix(grdMain.Rows - 1, 3) = Round((((rs.Fields("Selling_Price") / ((100 + rs.Fields("Sales_Tax")) / 100)) - rs.Fields("Landed_Cost")) / (rs.Fields("Selling_Price") / ((100 + rs.Fields("Sales_Tax")) / 100))) * 100, 3)
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Selling_Price") & "", "0.00")
        If Val(rs.Fields("Price2") & "") = 0 Then
            grdMain.TextMatrix(grdMain.Rows - 1, 5) = "0"
        Else
            grdMain.TextMatrix(grdMain.Rows - 1, 5) = Round((((rs.Fields("Price2") / ((100 + rs.Fields("Sales_Tax")) / 100)) - (rs.Fields("Selling_Price") / ((100 + rs.Fields("Sales_Tax")) / 100))) / (rs.Fields("Price2") / ((100 + rs.Fields("Sales_Tax")) / 100))) * 100, 3)
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Price2") & "", "0.00")
        rs.MoveNext
    Wend
    rs.Close
    grdTotal.MergeRow(0) = True
End Sub
Private Sub PreSales_Analysis()
    picMain.Visible = True
    cmb2.Width = 3225
    cmb1.Width = 2925
    cmb2.Left = 7380
    cmb1.Left = 10680
    cmb3.Visible = False
    grdMain.Cols = 7
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Product Code"
    grdMain.TextMatrix(0, 1) = " Description"
    grdMain.TextMatrix(0, 2) = " Landed Cost "
    grdMain.TextMatrix(0, 3) = " GP% "
    grdMain.TextMatrix(0, 4) = " Wholesale Price "
    grdMain.TextMatrix(0, 5) = " GP% "
    grdMain.TextMatrix(0, 6) = " Retail Price "
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignRightCenter
    grdMain.ColAlignment(3) = flexAlignRightCenter
    grdMain.ColAlignment(4) = flexAlignRightCenter
    grdMain.ColAlignment(5) = flexAlignRightCenter
    grdMain.ColAlignment(6) = flexAlignRightCenter
    grdMain.ColWidth(0) = grdMain.Width * 0.12
    grdMain.ColWidth(1) = grdMain.Width * 0.26
    grdMain.ColWidth(2) = grdMain.Width * 0.12
    grdMain.ColWidth(3) = grdMain.Width * 0.12
    grdMain.ColWidth(4) = grdMain.Width * 0.12
    grdMain.ColWidth(5) = grdMain.Width * 0.12
    grdMain.ColWidth(6) = grdMain.Width * 0.12
    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    lblCaption.Caption = "Reports - Pre-Sales Analysis"
    
    cmb1.Clear
    ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
    cmb1.AddItem "<All Departments>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<All Departments>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub User_Shift_List_Cols()
    picMain.Visible = True
    cmb2.Width = 3225
    cmb1.Width = 2925
    cmb2.Left = 7380
    cmb1.Left = 10680
    cmb3.Visible = False
    grdMain.Cols = 6
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Date & Time"
    grdMain.TextMatrix(0, 1) = "User"
    grdMain.TextMatrix(0, 2) = "Shift Start"
    grdMain.TextMatrix(0, 3) = "Shift Stop"
    grdMain.TextMatrix(0, 4) = "Commision"
    grdMain.TextMatrix(0, 5) = "Shift Duration"
    
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignLeftCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignRightCenter
    grdMain.ColAlignment(5) = flexAlignRightCenter
    
    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.31
    grdMain.ColWidth(2) = grdMain.Width * 0.15
    grdMain.ColWidth(3) = grdMain.Width * 0.15
    grdMain.ColWidth(4) = grdMain.Width * 0.1
    grdMain.ColWidth(5) = grdMain.Width * 0.1

    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    lblCaption.Caption = "Reports - User Shift List"
    ActiveReadServer "Select user_No,User_Name from Users order by user_no"
    cmb1.AddItem "<All Users>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<All Users>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Sales_Analysis_Hour()
    grdMain.FrozenCols = 1
    grdMain.ScrollBars = flexScrollBarBoth
    grdTotal.FrozenCols = 1
    picLive.Visible = False
    picMain.Visible = True
    cmb1.Width = 1845
    cmb2.Width = 2415
    cmb3.Width = 1905
    cmb1.Left = 9840
    cmb2.Left = 7380
    cmb3.Left = 11730
    cmb3.Visible = True
    grdMain.Cols = 26
    grdMain.FixedCols = 0
    grdMain.ColHidden(25) = True
    grdMain.TextMatrix(0, 0) = " Department"
    grdMain.TextMatrix(0, 1) = "07:00-08:00"
    grdMain.TextMatrix(0, 2) = "08:00-09:00"
    grdMain.TextMatrix(0, 3) = "09:00-10:00"
    grdMain.TextMatrix(0, 4) = "10:00-11:00"
    grdMain.TextMatrix(0, 5) = "11:00-12:00"
    grdMain.TextMatrix(0, 6) = "12:00-13:00"
    grdMain.TextMatrix(0, 7) = "13:00-14:00"
    grdMain.TextMatrix(0, 8) = "14:00-15:00"
    grdMain.TextMatrix(0, 9) = "15:00-16:00"
    grdMain.TextMatrix(0, 10) = "16:00-17:00"
    grdMain.TextMatrix(0, 11) = "17:00-18:00"
    grdMain.TextMatrix(0, 12) = "18:00-19:00"
    grdMain.TextMatrix(0, 13) = "19:00-20:00"
    grdMain.TextMatrix(0, 14) = "20:00-21:00"
    grdMain.TextMatrix(0, 15) = "21:00-22:00"
    grdMain.TextMatrix(0, 16) = "22:00-23:00"
    grdMain.TextMatrix(0, 17) = "23:00-24:00"
    grdMain.TextMatrix(0, 18) = "24:00-01:00"
    grdMain.TextMatrix(0, 19) = "01:00-02:00"
    grdMain.TextMatrix(0, 20) = "02:00-03:00"
    grdMain.TextMatrix(0, 21) = "03:00-04:00"
    grdMain.TextMatrix(0, 22) = "04:00-05:00"
    grdMain.TextMatrix(0, 23) = "05:00-06:00"
    grdMain.TextMatrix(0, 24) = "06:00-07:00"
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColWidth(0) = 3000
    For i = 1 To 24
        grdMain.ColWidth(i) = 950
        grdMain.ColAlignment(i) = flexAlignRightCenter
    Next i
    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    lblCaption.Caption = "Reports - Product Analysis"
    
    cmb1.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    cmb1.AddItem "<All Locations>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<All Locations>"
    
    cmb3.Clear
    ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
    cmb3.AddItem "<All Departments>"
    While Not rs.EOF
        cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb3.Text = "<All Departments>"
    grdMain.Tag = ""
    grdMain.Rows = 1
    Selection_Change
End Sub
Private Sub Placed_Purchaces()
    frmMain.Toolbar1.Buttons(5).Enabled = True
    picMain.Visible = True
    cmb1.Width = 1845
    cmb2.Width = 2415
    cmb3.Width = 1905
    cmb1.Left = 9840
    cmb2.Left = 7380
    cmb3.Left = 11730
    cmb3.Visible = True
    grdMain.Cols = 5
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Order Date"
    grdMain.TextMatrix(0, 1) = " Account"
    grdMain.TextMatrix(0, 2) = " User "
    grdMain.TextMatrix(0, 3) = " Description "
    grdMain.TextMatrix(0, 4) = " Total (incl) "
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignLeftCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignRightCenter
    grdMain.ColWidth(0) = grdMain.Width * 0.17
    grdMain.ColWidth(1) = grdMain.Width * 0.3
    grdMain.ColWidth(2) = grdMain.Width * 0.17
    grdMain.ColWidth(3) = grdMain.Width * 0.2
    grdMain.ColWidth(4) = grdMain.Width * 0.12
    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    lblCaption.Caption = "Reports - Placed Orders"
    
    cmb1.Clear
    ActiveReadServer "Select Supplier_No,Supplier_Name from suppliers order by Supplier_name"
    cmb1.AddItem "<All Suppliers>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<All Suppliers>"
    
    cmb3.Clear
    ActiveReadServer "Select user_No,User_Name from Users order by user_no"
    cmb3.AddItem "<All Users>"
    While Not rs.EOF
        cmb3.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
        rs.MoveNext
    Wend
    cmb3.Text = "<All Users>"
    rs.Close
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Debtor_Accounts()
    picMain.Visible = True
    cmb2.Width = 3225
    cmb1.Width = 2925
    cmb2.Left = 7380
    cmb1.Left = 10680
    cmb3.Visible = False
    grdMain.Cols = 6
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Debtor No"
    grdMain.TextMatrix(0, 1) = "Debtor Name"
    grdMain.TextMatrix(0, 2) = "Contact Person"
    grdMain.TextMatrix(0, 3) = "Telephone No."
    grdMain.TextMatrix(0, 4) = "Type"
    grdMain.TextMatrix(0, 5) = "Balance"
    
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignLeftCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignLeftCenter
    grdMain.ColAlignment(5) = flexAlignRightCenter
    
    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.25
    grdMain.ColWidth(2) = grdMain.Width * 0.15
    grdMain.ColWidth(3) = grdMain.Width * 0.15
    grdMain.ColWidth(4) = grdMain.Width * 0.15
    grdMain.ColWidth(5) = grdMain.Width * 0.1

    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    lblCaption.Caption = "Reports - Debtors List"
    cmb1.AddItem "<All Debtors>"
    cmb1.AddItem "Debtor Accounts"
    cmb1.AddItem "Staff Accounts"
    cmb1.AddItem "Management Accounts"
    cmb1.AddItem "Travel Agents"
    cmb1.AddItem "Members"
    cmb1.Text = "<All Debtors>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Debtor_Age_Analysis()
    picMain.Visible = True
    cmb1.Width = 2245
    cmb2.Width = 1800
    cmb3.Width = 2100
    cmb1.Left = 9240
    cmb2.Left = 7380
    cmb3.Left = 11530
    cmb3.Visible = True
    grdMain.Cols = 8
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Debtor No"
    grdMain.TextMatrix(0, 1) = "Debtor Name"
    grdMain.TextMatrix(0, 2) = "120 Days+"
    grdMain.TextMatrix(0, 3) = "90 Days"
    grdMain.TextMatrix(0, 4) = "60 Days"
    grdMain.TextMatrix(0, 5) = "30 Days"
    grdMain.TextMatrix(0, 6) = "Current"
    grdMain.TextMatrix(0, 7) = "Balance"
    
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignRightCenter
    grdMain.ColAlignment(3) = flexAlignRightCenter
    grdMain.ColAlignment(4) = flexAlignRightCenter
    grdMain.ColAlignment(5) = flexAlignRightCenter
    grdMain.ColAlignment(6) = flexAlignRightCenter
    grdMain.ColAlignment(7) = flexAlignRightCenter
    
    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.25
    grdMain.ColWidth(2) = grdMain.Width * 0.1
    grdMain.ColWidth(3) = grdMain.Width * 0.1
    grdMain.ColWidth(4) = grdMain.Width * 0.1
    grdMain.ColWidth(5) = grdMain.Width * 0.1
    grdMain.ColWidth(6) = grdMain.Width * 0.1
    grdMain.ColWidth(7) = grdMain.Width * 0.1

    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    lblCaption.Caption = "Reports - Debtors List"
    cmb1.AddItem "<All Debtors>"
    cmb1.AddItem "Debtor Accounts"
    cmb1.AddItem "Staff Accounts"
    cmb1.AddItem "Management Accounts"
    cmb1.AddItem "Travel Agents"
    cmb1.AddItem "Members"
    cmb1.Text = "<All Debtors>"
    cmb3.Clear
    lblCaption.Caption = "Reports - Debtors Age Analysis"
    cmb3.AddItem "<All Balances>"
    cmb3.AddItem "120 Days+"
    cmb3.AddItem "90 Days"
    cmb3.AddItem "60 Days"
    cmb3.AddItem "30 Days"
    cmb3.AddItem "Current"
    cmb3.Text = "<All Balances>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Cost_Price_Variations()
    picMain.Visible = True
    cmdSearch.Visible = True
    cmdPrice.Visible = True
    cmb1.Width = 2245
    cmb2.Width = 1800
    cmb3.Width = 2100
    cmb1.Left = 9240
    cmb2.Left = 7380
    cmb3.Left = 11530
    cmb3.Visible = True
    grdMain.Cols = 7
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Product Code"
    grdMain.TextMatrix(0, 1) = "Description"
    grdMain.TextMatrix(0, 2) = "Supplier"
    grdMain.TextMatrix(0, 3) = "Invoice No"
    grdMain.TextMatrix(0, 4) = "Old Cost Price"
    grdMain.TextMatrix(0, 5) = "Variance %"
    grdMain.TextMatrix(0, 6) = "New Cost Price"
    
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignLeftCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignRightCenter
    grdMain.ColAlignment(5) = flexAlignRightCenter
    grdMain.ColAlignment(6) = flexAlignRightCenter
    
    grdMain.ColWidth(0) = grdMain.Width * 0.12
    grdMain.ColWidth(1) = grdMain.Width * 0.23
    grdMain.ColWidth(2) = grdMain.Width * 0.25
    grdMain.ColWidth(3) = grdMain.Width * 0.1
    grdMain.ColWidth(4) = grdMain.Width * 0.1
    grdMain.ColWidth(5) = grdMain.Width * 0.1
    grdMain.ColWidth(6) = grdMain.Width * 0.1

    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    lblCaption.Caption = "Reports - Cost Price Variations"
    cmb1.Clear
    ActiveReadServer "Select Supplier_No,Supplier_Name from suppliers order by Supplier_name"
    cmb1.AddItem "<All Suppliers>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<All Suppliers>"
    cmb3.Clear
    ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
    cmb3.AddItem "<All Departments>"
    While Not rs.EOF
        cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb3.Text = "<All Departments>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Table_Tranfer_Journal()
    picMain.Visible = True
    cmb1.Width = 2245
    cmb2.Width = 1800
    cmb3.Width = 2100
    cmb1.Left = 9240
    cmb2.Left = 7380
    cmb3.Left = 11530
    cmb3.Visible = True
    grdMain.Cols = 9
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Date & Time"
    grdMain.TextMatrix(0, 1) = "Transfering User"
    grdMain.TextMatrix(0, 2) = "Receiving User"
    grdMain.TextMatrix(0, 3) = "From Table"
    grdMain.TextMatrix(0, 4) = "To Table"
    grdMain.TextMatrix(0, 5) = "From Tab"
    grdMain.TextMatrix(0, 6) = "To Tab"
    grdMain.TextMatrix(0, 7) = "Invoice No"
    grdMain.TextMatrix(0, 8) = "Action"
    
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignLeftCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignLeftCenter
    grdMain.ColAlignment(5) = flexAlignLeftCenter
    grdMain.ColAlignment(6) = flexAlignLeftCenter
    grdMain.ColAlignment(7) = flexAlignLeftCenter
    grdMain.ColAlignment(8) = flexAlignLeftCenter
    
    grdMain.ColWidth(0) = grdMain.Width * 0.11
    grdMain.ColWidth(1) = grdMain.Width * 0.15
    grdMain.ColWidth(2) = grdMain.Width * 0.15
    grdMain.ColWidth(3) = grdMain.Width * 0.08
    grdMain.ColWidth(4) = grdMain.Width * 0.08
    grdMain.ColWidth(5) = grdMain.Width * 0.1
    grdMain.ColWidth(6) = grdMain.Width * 0.1
    grdMain.ColWidth(7) = grdMain.Width * 0.08
    grdMain.ColWidth(8) = grdMain.Width * 0.15
    
    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    cmb3.Clear
    lblCaption.Caption = "Reports - Table and Tab Journal"
    ActiveReadServer "Select user_No,User_Name from Users order by user_no"
    cmb1.AddItem "<All Users>"
    cmb3.AddItem "<All Users>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
        cmb3.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<All Users>"
    cmb3.Text = "<All Users>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Staff_Wages_Screen()
    picMain.Visible = True
    cmb2.Width = 3225
    cmb1.Width = 2925
    cmb2.Left = 7380
    cmb1.Left = 10680
    cmb3.Visible = False
    grdMain.Cols = 6
    grdMain.FixedCols = 0
    grdMain.TextMatrix(0, 0) = " Date & Time"
    grdMain.TextMatrix(0, 1) = "User"
    grdMain.TextMatrix(0, 2) = "Shift Start"
    grdMain.TextMatrix(0, 3) = "Shift Stop"
    grdMain.TextMatrix(0, 4) = "Commision"
    grdMain.TextMatrix(0, 5) = "Shift Duration"
    
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignLeftCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignRightCenter
    grdMain.ColAlignment(5) = flexAlignRightCenter
    
    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.31
    grdMain.ColWidth(2) = grdMain.Width * 0.15
    grdMain.ColWidth(3) = grdMain.Width * 0.15
    grdMain.ColWidth(4) = grdMain.Width * 0.1
    grdMain.ColWidth(5) = grdMain.Width * 0.1

    grdTotal.Cols = grdMain.Cols
    For i = 0 To grdMain.Cols - 1
        grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
        grdTotal.ColWidth(i) = grdMain.ColWidth(i)
    Next i
    grdMain.Rows = 1
    grdMain.Tag = "1"
    cmb1.Clear
    lblCaption.Caption = "Reports - User Shift List"
    ActiveReadServer "Select user_No,User_Name from Users order by user_no"
    cmb1.AddItem "<All Users>"
    While Not rs.EOF
        cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb1.Text = "<All Users>"
    grdMain.Tag = ""
    Selection_Change
End Sub
Private Sub Stocklevelsarelow() 'Stock Levels Low
picBlocDate.Visible = True
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 9
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Code"
            grdMain.TextMatrix(0, 1) = " Description"
            grdMain.TextMatrix(0, 2) = " Unit Size"
            grdMain.TextMatrix(0, 3) = " Department"
            grdMain.TextMatrix(0, 4) = "SOH "
            grdMain.TextMatrix(0, 5) = "Unit Cost "
            grdMain.TextMatrix(0, 6) = "Total (excl) "
            grdMain.TextMatrix(0, 7) = "Tax "
            grdMain.TextMatrix(0, 8) = "Total (incl) "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColAlignment(8) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.1
            grdMain.ColWidth(1) = grdMain.Width * 0.29
            grdMain.ColWidth(2) = grdMain.Width * 0.06
            grdMain.ColWidth(3) = grdMain.Width * 0.15
            grdMain.ColWidth(4) = grdMain.Width * 0.08
            grdMain.ColWidth(5) = grdMain.Width * 0.08
            grdMain.ColWidth(6) = grdMain.Width * 0.08
            grdMain.ColWidth(7) = grdMain.Width * 0.08
            grdMain.ColWidth(8) = grdMain.Width * 0.08
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Stock Levels Low"
            
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Locations>"
            
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            grdMain.Tag = ""
            Selection_Change
End Sub




Private Sub cmb2_Change()
Screen.MousePointer = vbHourglass
Buttime.Visible = False
    Lbltime.Visible = False
    Lbltimeback.Visible = False
    On Error Resume Next
    frmMain.Toolbar1.Buttons(5).Enabled = False
    grdTotal.Visible = True
    grdMain.Height = 7710
    picLive.Visible = False
    grdMain.ScrollBars = flexScrollBarVertical
    grdTotal.FrozenCols = 0
    grdTotal.ScrollBars = flexScrollBarNone
    grdMain.FrozenCols = 0
    Screen.MousePointer = 11
    DoEvents
    ButtonEx1.Visible = True
    picAnalysis.Visible = False
    picMain.Visible = False
    grdMain.Rows = 1
    picTrade.Visible = False
    grdMain.ColHidden(0) = False
    grdTotal.ColHidden(0) = False
    grdMain.RowHeight(0) = 350
    grdMain.Cols = 2
    picBlocDate.Visible = False
    If Screen.MousePointer <> 11 Then Screen.MousePointer = 11
    Select Case cmb2.Text
        Case "Staff Wages"
            Staff_Wages_Screen
        Case "Pack Links"
            Pack_Links_Screen
        Case "Table and Tab Journal"
            Table_Tranfer_Journal
        Case "Cost Price Variations"
            Cost_Price_Variations
        Case "Age Analysis"
            picBlocDate.Visible = True
            Debtor_Age_Analysis
        Case "Debtor Accounts"
            picBlocDate.Visible = True
            Debtor_Accounts
        Case "Placed Purchase Orders"
            Placed_Purchaces
        Case "Payout Journal"
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 5
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Date & Time"
            grdMain.TextMatrix(0, 1) = " Account"
            grdMain.TextMatrix(0, 2) = " User "
            grdMain.TextMatrix(0, 3) = " Description "
            grdMain.TextMatrix(0, 4) = " Total (incl) "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.17
            grdMain.ColWidth(1) = grdMain.Width * 0.3
            grdMain.ColWidth(2) = grdMain.Width * 0.17
            grdMain.ColWidth(3) = grdMain.Width * 0.28
            grdMain.ColWidth(4) = grdMain.Width * 0.07
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Payout Journal"
            
            cmb1.Clear
            ActiveReadServer "Select Supplier_No,Supplier_Name from suppliers order by Supplier_name"
            cmb1.AddItem "<All Suppliers>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Suppliers>"
            cmb3.Clear
            ActiveReadServer "Select user_No,User_Name from Users order by user_no"
            cmb3.AddItem "<All Users>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.MoveNext
            Wend
            cmb3.Text = "<All Users>"
            rs.Close
            grdMain.Tag = ""
            Selection_Change
        Case "Pre-Sales Analysis by Price"
            picBlocDate.Visible = True
            PreSales_Analysis
        Case "Stock Summary"
            picBlocDate.Visible = True
            picMain.Visible = True
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            grdMain.Cols = 7
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Department No"
            grdMain.TextMatrix(0, 1) = " Department"
            grdMain.TextMatrix(0, 2) = " Stock on Hand "
            grdMain.TextMatrix(0, 3) = " Ave Cost per Unit "
            grdMain.TextMatrix(0, 4) = " Total (excl) "
            grdMain.TextMatrix(0, 5) = " Tax "
            grdMain.TextMatrix(0, 6) = " Total (incl) "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignRightCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.12
            grdMain.ColWidth(1) = grdMain.Width * 0.26
            grdMain.ColWidth(2) = grdMain.Width * 0.12
            grdMain.ColWidth(3) = grdMain.Width * 0.12
            grdMain.ColWidth(4) = grdMain.Width * 0.12
            grdMain.ColWidth(5) = grdMain.Width * 0.12
            grdMain.ColWidth(6) = grdMain.Width * 0.12
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Stock on Hand"
            
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Locations>"
            
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            grdMain.Tag = ""
            Selection_Change
        Case "Stock on Hand (Suppliers)"
            picBlocDate.Visible = True
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 9
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Code"
            grdMain.TextMatrix(0, 1) = " Description"
            grdMain.TextMatrix(0, 2) = " Unit Size"
            grdMain.TextMatrix(0, 3) = " Supplier"
            grdMain.TextMatrix(0, 4) = "SOH "
            grdMain.TextMatrix(0, 5) = "Unit Cost "
            grdMain.TextMatrix(0, 6) = "Total (excl) "
            grdMain.TextMatrix(0, 7) = "Tax "
            grdMain.TextMatrix(0, 8) = "Total (incl) "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColAlignment(8) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.08
            grdMain.ColWidth(1) = grdMain.Width * 0.2
            grdMain.ColWidth(2) = grdMain.Width * 0.1
            grdMain.ColWidth(3) = grdMain.Width * 0.24
            grdMain.ColWidth(4) = grdMain.Width * 0.07
            grdMain.ColWidth(5) = grdMain.Width * 0.07
            grdMain.ColWidth(6) = grdMain.Width * 0.07
            grdMain.ColWidth(7) = grdMain.Width * 0.07
            grdMain.ColWidth(8) = grdMain.Width * 0.08
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Stock on Hand"
            cmb1.Clear
            ActiveReadServer "Select Supplier_No,Supplier_Name from suppliers order by Supplier_name"
            cmb1.AddItem "<All Suppliers>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Suppliers>"
            
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            grdMain.Tag = ""
            Selection_Change
        
        
        
        
        Case "Stock Levels Low"  'Stocklow_View
        Call Stocklevelsarelow
        
        
        Case "Sales Analysis by Hour"
            Sales_Analysis_Hour
        Case "Product Analysis"
            picBlocDate.Visible = True
            picMain.Visible = True
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            grdMain.Cols = 7
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Code"
            grdMain.TextMatrix(0, 1) = " Description"
            grdMain.TextMatrix(0, 2) = " Department"
            grdMain.TextMatrix(0, 3) = " Unit Cost "
            grdMain.TextMatrix(0, 4) = " Markup%"
            grdMain.TextMatrix(0, 5) = " GP%"
            grdMain.TextMatrix(0, 6) = " Selling (Incl) "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.25
            grdMain.ColWidth(2) = grdMain.Width * 0.2
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.1
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Product Analysis"
            
            cmb1.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb1.AddItem "<All Departments>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Departments>"
            grdMain.Tag = ""
            Selection_Change
         Case "Stock on Hand"
            picBlocDate.Visible = True
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 9
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Code"
            grdMain.TextMatrix(0, 1) = " Description"
            grdMain.TextMatrix(0, 2) = " Unit Size"
            grdMain.TextMatrix(0, 3) = " Department"
            grdMain.TextMatrix(0, 4) = "SOH "
            grdMain.TextMatrix(0, 5) = "Unit Cost "
            grdMain.TextMatrix(0, 6) = "Total (excl) "
            grdMain.TextMatrix(0, 7) = "Tax "
            grdMain.TextMatrix(0, 8) = "Total (incl) "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColAlignment(8) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.1
            grdMain.ColWidth(1) = grdMain.Width * 0.29
            grdMain.ColWidth(2) = grdMain.Width * 0.06
            grdMain.ColWidth(3) = grdMain.Width * 0.15
            grdMain.ColWidth(4) = grdMain.Width * 0.08
            grdMain.ColWidth(5) = grdMain.Width * 0.08
            grdMain.ColWidth(6) = grdMain.Width * 0.08
            grdMain.ColWidth(7) = grdMain.Width * 0.08
            grdMain.ColWidth(8) = grdMain.Width * 0.08
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Stock on Hand"
            
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Locations>"
            
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            grdMain.Tag = ""
            Selection_Change
        Case "Analysis by Department"
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            grdMain.Cols = 4
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Date & Time"
            grdMain.TextMatrix(0, 1) = "Department"
            grdMain.TextMatrix(0, 2) = ""
            grdMain.TextMatrix(0, 3) = "Transaction Type"
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.2
            grdMain.ColWidth(1) = grdMain.Width * 0.3
            grdMain.ColWidth(2) = grdMain.Width * 0.2
            grdMain.ColWidth(3) = grdMain.Width * 0.3
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Purchase Analysis"
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb1.AddItem "<All Departments>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Departments>"
            grdMain.Tag = ""
            Selection_Change
        Case "Stock Takes"
            grdMain.RowHeight(0) = 350
            If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            grdMain.Cols = 7
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Date & Time"
            grdMain.TextMatrix(0, 1) = "Take No"
            grdMain.TextMatrix(0, 2) = "Location"
            grdMain.TextMatrix(0, 3) = "Transaction Type"
            grdMain.TextMatrix(0, 4) = "Stock on Hand" '
            grdMain.TextMatrix(0, 5) = "Variance"
            grdMain.TextMatrix(0, 6) = "Counted"
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.2
            grdMain.ColWidth(1) = grdMain.Width * 0.1
            grdMain.ColWidth(2) = grdMain.Width * 0.2
            grdMain.ColWidth(3) = grdMain.Width * 0.2
            grdMain.ColWidth(4) = grdMain.Width * 0.1
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
'            Stocktakeupdates
'             cmb1.AddItem "<Extra Stock Functions>"
'             cmb1.AddItem "Stock not counted during stocktake"
            
            cmb1.Text = "<All Locations>"
            grdMain.Tag = ""
            Selection_Change
        Case "Discounted Sales"
            frmMain.Toolbar1.Buttons(14).Enabled = True
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            lblCaption.Caption = "Reports - Discounted Sales"
            grdMain.Cols = 7
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Date & Time"
            grdMain.TextMatrix(0, 1) = " Product Code"
            grdMain.TextMatrix(0, 2) = "Description"
            grdMain.TextMatrix(0, 3) = "Invoice No"
            grdMain.TextMatrix(0, 4) = "Qty Corrected "
            grdMain.TextMatrix(0, 5) = "Line_Total "
            grdMain.TextMatrix(0, 6) = " User"
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignLeftCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.1
            grdMain.ColWidth(2) = grdMain.Width * 0.25
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.1
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            ActiveReadServer "Select user_No,User_Name from Users order by user_no"
            cmb1.AddItem "<All Users>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.MoveNext
            Wend
            cmb1.Text = "<All Users>"
            rs.Close
            grdMain.Tag = ""
            Selection_Change
        Case "Sales Corrections"
            frmMain.Toolbar1.Buttons(14).Enabled = True
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            lblCaption.Caption = "Reports - Sales Corrections"
            grdMain.Cols = 7
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Date & Time"
            grdMain.TextMatrix(0, 1) = " Product Code"
            grdMain.TextMatrix(0, 2) = "Description"
            grdMain.TextMatrix(0, 3) = "Invoice No"
            grdMain.TextMatrix(0, 4) = "Qty Corrected "
            grdMain.TextMatrix(0, 5) = "Line_Total "
            grdMain.TextMatrix(0, 6) = " User"
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignLeftCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.1
            grdMain.ColWidth(2) = grdMain.Width * 0.25
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.1
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            cmb1.AddItem "<All Corrections>"
            cmb1.AddItem "Item Corrects"
            cmb1.AddItem "Voids"
            cmb1.AddItem "Returns"
            cmb1.Text = "<All Corrections>"
            cmb3.Clear
            ActiveReadServer "Select user_No,User_Name from Users order by user_no"
            cmb3.AddItem "<All Users>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.MoveNext
            Wend
            cmb3.Text = "<All Users>"
            rs.Close
            grdMain.Tag = ""
            Selection_Change
        Case "Staff Commision Report"
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            grdMain.Cols = 10
            grdMain.FixedCols = 0
            grdMain.RowHeight(0) = 650
            grdMain.TextMatrix(0, 0) = "User"
            grdMain.TextMatrix(0, 1) = "Cash "
            grdMain.TextMatrix(0, 2) = "Card "
            grdMain.TextMatrix(0, 3) = "Voucher "
            grdMain.TextMatrix(0, 4) = "Charge "
            grdMain.TextMatrix(0, 5) = "Total (Incl) "
            grdMain.TextMatrix(0, 6) = "Tax "
            grdMain.TextMatrix(0, 7) = "Total (Excl) "
            grdMain.TextMatrix(0, 8) = "Comm% "
            grdMain.TextMatrix(0, 9) = "Commision Due "
            
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignRightCenter
            grdMain.ColAlignment(2) = flexAlignRightCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColAlignment(8) = flexAlignRightCenter
            grdMain.ColAlignment(9) = flexAlignRightCenter
            
            grdMain.ColWidth(0) = grdMain.Width * 0.18
            grdMain.ColWidth(1) = grdMain.Width * 0.09
            grdMain.ColWidth(2) = grdMain.Width * 0.09
            grdMain.ColWidth(3) = grdMain.Width * 0.09
            grdMain.ColWidth(4) = grdMain.Width * 0.09
            grdMain.ColWidth(5) = grdMain.Width * 0.09
            grdMain.ColWidth(6) = grdMain.Width * 0.09
            grdMain.ColWidth(7) = grdMain.Width * 0.09
            grdMain.ColWidth(8) = grdMain.Width * 0.09
            grdMain.ColWidth(9) = grdMain.Width * 0.09

            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - User Shift List"
            ActiveReadServer "Select user_No,User_Name from Users order by user_no"
            cmb1.AddItem "<All Users>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Users>"
            grdMain.Tag = ""
            Selection_Change
        Case "Staff Shift Report"
            User_Shift_List_Cols
        Case "Room Accounts"
            grdMain.RowHeight(0) = 450
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            picMain.Visible = True
            grdMain.Cols = 8
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Room"
            grdMain.TextMatrix(0, 1) = "Guest Name"
            grdMain.TextMatrix(0, 2) = "Arrival Date"
            grdMain.TextMatrix(0, 3) = "Departure Date"
            grdMain.TextMatrix(0, 4) = "Status"
            grdMain.TextMatrix(0, 5) = "Nights"
            grdMain.TextMatrix(0, 6) = "Balance"
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColHidden(7) = True
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.25
            grdMain.ColWidth(2) = grdMain.Width * 0.15
            grdMain.ColWidth(3) = grdMain.Width * 0.15
            grdMain.ColWidth(4) = grdMain.Width * 0.12
            grdMain.ColWidth(5) = grdMain.Width * 0.09
            grdMain.ColWidth(6) = grdMain.Width * 0.09
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            cmb1.AddItem "<All Reservations>"
            cmb1.AddItem "Confirmed Bookings"
            cmb1.AddItem "Guests Checked In"
            cmb1.AddItem "Guests Checked Out"
            cmb1.Text = "<All Reservations>"
            cmb3.Clear
            lblCaption.Caption = "Reports - Room Accounts"
            grdMain.Tag = ""
            Selection_Change
        Case "Room Sales"
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 8
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Date & Time"
            grdMain.TextMatrix(0, 1) = "User"
            grdMain.TextMatrix(0, 2) = "Workstation"
            grdMain.TextMatrix(0, 3) = "Document No."
            grdMain.TextMatrix(0, 4) = "Transaction Type"
            grdMain.TextMatrix(0, 5) = "Total (excl) "
            grdMain.TextMatrix(0, 6) = "Tax "
            grdMain.TextMatrix(0, 7) = "Total (incl) "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignLeftCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.15
            grdMain.ColWidth(2) = grdMain.Width * 0.15
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.15
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdMain.ColWidth(7) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Sales Journal"
            ActiveReadServer "Select user_No,User_Name from Users order by user_no"
            cmb1.AddItem "<All Users>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Users>"
            cmb3.Clear
            cmb3.AddItem "<All Transactions>"
            cmb3.AddItem "Room Accomodation"
            cmb3.AddItem "Other"
            cmb3.Text = "<All Transactions>"
            grdMain.Tag = ""
            Selection_Change
        Case "Receive on Account"
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            picMain.Visible = True
            picMain.Visible = True
            grdMain.Cols = 6
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Dated"
            grdMain.TextMatrix(0, 1) = "Debtor No"
            grdMain.TextMatrix(0, 2) = "Debtor Name"
            grdMain.TextMatrix(0, 3) = "Tel No"
            grdMain.TextMatrix(0, 4) = "Payment "
            grdMain.TextMatrix(0, 5) = "Payment Type"
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.12
            grdMain.ColWidth(1) = grdMain.Width * 0.15
            grdMain.ColWidth(2) = grdMain.Width * 0.32
            grdMain.ColWidth(3) = grdMain.Width * 0.12
            grdMain.ColWidth(4) = grdMain.Width * 0.12
            grdMain.ColWidth(5) = grdMain.Width * 0.09
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            cmb1.AddItem "<All Tender Types>"
            cmb1.AddItem "Cash"
            cmb1.AddItem "Card"
            cmb1.AddItem "Voucher"
            cmb1.AddItem "Charge"
            cmb1.AddItem "EFT"
            cmb1.Text = "<All Tender Types>"
            lblCaption.Caption = "Reports - Receive on Account"
            grdMain.Tag = ""
            Selection_Change
        Case "Deposits Paid", "Payments Received"
            grdMain.RowHeight(0) = 450
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            picMain.Visible = True
            picMain.Visible = True
            grdMain.Cols = 9
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Dated"
            grdMain.TextMatrix(0, 1) = " Room"
            grdMain.TextMatrix(0, 2) = "Guest Name"
            grdMain.TextMatrix(0, 3) = "Arrival Date"
            grdMain.TextMatrix(0, 4) = "Departure Date"
            grdMain.TextMatrix(0, 5) = "Status"
            grdMain.TextMatrix(0, 6) = "Deposit "
            grdMain.TextMatrix(0, 7) = "Payment Type"
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColHidden(8) = True
            grdMain.ColWidth(0) = grdMain.Width * 0.12
            grdMain.ColWidth(1) = grdMain.Width * 0.15
            grdMain.ColWidth(2) = grdMain.Width * 0.22
            grdMain.ColWidth(3) = grdMain.Width * 0.12
            grdMain.ColWidth(4) = grdMain.Width * 0.12
            grdMain.ColWidth(5) = grdMain.Width * 0.09
            grdMain.ColWidth(6) = grdMain.Width * 0.09
            grdMain.ColWidth(7) = grdMain.Width * 0.09
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            cmb1.AddItem "<All Reservations>"
            cmb1.AddItem "Confirmed Bookings"
            cmb1.AddItem "Guests Checked In"
            cmb1.AddItem "Guests Checked Out"
            cmb1.Text = "<All Reservations>"
            cmb3.Clear
            cmb3.AddItem "<All Tender Types>"
            cmb3.AddItem "Cash"
            cmb3.AddItem "Card"
            cmb3.AddItem "Voucher"
            cmb3.AddItem "Charge"
            cmb3.AddItem "EFT"
            cmb3.Text = "<All Tender Types>"
            lblCaption.Caption = "Reports - Deposits Received"
            If cmb2.Text = "Payments Received" Then
                lblCaption.Caption = "Reports - Payments Received"
                grdMain.TextMatrix(0, 5) = "Payment"
            End If
            grdMain.Tag = ""
            Selection_Change
        Case "Stock Movement (Values)"
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 13
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Product Code"
            grdMain.TextMatrix(0, 1) = "Description"
            grdMain.TextMatrix(0, 2) = "Opening Stock"
            grdMain.TextMatrix(0, 3) = "Goods Received"
            grdMain.TextMatrix(0, 4) = "Goods Returned"
            grdMain.TextMatrix(0, 5) = "Wastage"
            grdMain.TextMatrix(0, 6) = "Outgoing Transfers"
            grdMain.TextMatrix(0, 7) = "Incoming Transfers"
            grdMain.TextMatrix(0, 8) = "Production"
            grdMain.TextMatrix(0, 9) = "Sales Consumption"
            grdMain.TextMatrix(0, 10) = "Theoretical Close"
            grdMain.TextMatrix(0, 11) = "Variance"
            grdMain.TextMatrix(0, 12) = "Closing Stock"
            grdMain.RowHeight(0) = 650
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignRightCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColAlignment(8) = flexAlignRightCenter
            grdMain.ColAlignment(9) = flexAlignRightCenter
            grdMain.ColAlignment(10) = flexAlignRightCenter
            grdMain.ColAlignment(11) = flexAlignRightCenter
            grdMain.ColAlignment(12) = flexAlignRightCenter
            grdMain.ColHidden(0) = True
            grdTotal.ColHidden(0) = True
            grdMain.ColWidth(1) = grdMain.Width * 0.21
            grdMain.ColWidth(2) = grdMain.Width * 0.07
            grdMain.ColWidth(3) = grdMain.Width * 0.07
            grdMain.ColWidth(4) = grdMain.Width * 0.07
            grdMain.ColWidth(5) = grdMain.Width * 0.07
            grdMain.ColWidth(6) = grdMain.Width * 0.06
            grdMain.ColWidth(7) = grdMain.Width * 0.06
            grdMain.ColWidth(8) = grdMain.Width * 0.07
            grdMain.ColWidth(9) = grdMain.Width * 0.073
            grdMain.ColWidth(10) = grdMain.Width * 0.07
            grdMain.ColWidth(11) = grdMain.Width * 0.07
            grdMain.ColWidth(12) = grdMain.Width * 0.07
            
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Locations>"
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            lblCaption.Caption = "Reports - Stock Movement (Values)"
            grdMain.Tag = ""
            Selection_Change
        Case "Stock Movement (Quantities)"
            DoEvents
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 13
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Product Code"
            grdMain.TextMatrix(0, 1) = "Description"
            grdMain.TextMatrix(0, 2) = "Opening Stock"
            grdMain.TextMatrix(0, 3) = "Goods Received"
            grdMain.TextMatrix(0, 4) = "Goods Returned"
            grdMain.TextMatrix(0, 5) = "Wastage"
            grdMain.TextMatrix(0, 6) = "Outgoing Transfers"
            grdMain.TextMatrix(0, 7) = "Incoming Transfers"
            grdMain.TextMatrix(0, 8) = "Production"
            grdMain.TextMatrix(0, 9) = "Sales Consumption"
            grdMain.TextMatrix(0, 10) = "Theoretical Close"
            grdMain.TextMatrix(0, 11) = "Variance"
            grdMain.TextMatrix(0, 12) = "Closing Stock"
            grdMain.RowHeight(0) = 650
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignRightCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColAlignment(8) = flexAlignRightCenter
            grdMain.ColAlignment(9) = flexAlignRightCenter
            grdMain.ColAlignment(10) = flexAlignRightCenter
            grdMain.ColAlignment(11) = flexAlignRightCenter
            grdMain.ColAlignment(12) = flexAlignRightCenter
            grdMain.ColHidden(0) = True
            grdTotal.ColHidden(0) = True
            grdMain.ColWidth(1) = grdMain.Width * 0.21
            grdMain.ColWidth(2) = grdMain.Width * 0.07
            grdMain.ColWidth(3) = grdMain.Width * 0.07
            grdMain.ColWidth(4) = grdMain.Width * 0.07
            grdMain.ColWidth(5) = grdMain.Width * 0.07
            grdMain.ColWidth(6) = grdMain.Width * 0.06
            grdMain.ColWidth(7) = grdMain.Width * 0.06
            grdMain.ColWidth(8) = grdMain.Width * 0.07
            grdMain.ColWidth(9) = grdMain.Width * 0.073
            grdMain.ColWidth(10) = grdMain.Width * 0.07
            grdMain.ColWidth(11) = grdMain.Width * 0.07
            grdMain.ColWidth(12) = grdMain.Width * 0.07
            
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Locations>"
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            lblCaption.Caption = "Reports - Stock Movement (Quantities)"
            grdMain.Tag = ""
            Selection_Change
        Case "Stock Variance"
            grdMain.RowHeight(0) = 600
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 9
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Product Code"
            grdMain.TextMatrix(0, 1) = "Description"
            grdMain.TextMatrix(0, 2) = "Department"
            grdMain.TextMatrix(0, 3) = "Theoretical on Hand"
            grdMain.TextMatrix(0, 4) = "Variance"
            grdMain.TextMatrix(0, 5) = "Closing Stock"
            grdMain.TextMatrix(0, 6) = "Variance Value"
            grdMain.TextMatrix(0, 7) = "Stock Value"
            grdMain.ColHidden(8) = True
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.1
            grdMain.ColWidth(1) = grdMain.Width * 0.2
            grdMain.ColWidth(2) = grdMain.Width * 0.2
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.1
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdMain.ColWidth(7) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Tag = "1"
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Locations>"
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            lblCaption.Caption = "Reports - Stock Variance"
            grdMain.Tag = ""
            Selection_Change
        Case "User Journals"
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            grdMain.Cols = 4
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Date & Time"
            grdMain.TextMatrix(0, 1) = "User"
            grdMain.TextMatrix(0, 2) = "Workstation"
            grdMain.TextMatrix(0, 3) = "Transaction Type"
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.2
            grdMain.ColWidth(1) = grdMain.Width * 0.3
            grdMain.ColWidth(2) = grdMain.Width * 0.2
            grdMain.ColWidth(3) = grdMain.Width * 0.3
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - User Journals"
            ActiveReadServer "Select user_No,User_Name from Users order by user_no"
            cmb1.AddItem "<All Users>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Users>"
            grdMain.Tag = ""
            Selection_Change
        Case "Transfer Journal"
            Transfer_Journal
        Case "Sales Journal"
            ActiveUpdateServer "Delete from Sales_Temp"
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 9
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Date & Time"
            grdMain.TextMatrix(0, 1) = "User"
            grdMain.TextMatrix(0, 2) = "Detail"
            grdMain.TextMatrix(0, 3) = "Document No."
            grdMain.TextMatrix(0, 4) = "Transaction Type"
            grdMain.TextMatrix(0, 5) = "Total (excl) "
            grdMain.TextMatrix(0, 6) = "Tax "
            grdMain.TextMatrix(0, 7) = "Total (incl) "
            grdMain.TextMatrix(0, 8) = "Pax"
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignLeftCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColAlignment(8) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.15
            grdMain.ColWidth(2) = grdMain.Width * 0.12
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.12 '8
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdMain.ColWidth(7) = grdMain.Width * 0.1
            grdMain.ColWidth(8) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Sales Journal"
            ActiveReadServer "Select user_No,User_Name from Users order by user_no"
            cmb1.AddItem "<All Users>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Users>"
            cmb3.Clear
            cmb3.AddItem "<All Transactions>"
            cmb3.AddItem "Cash Sales"
            cmb3.AddItem "Card Sales"
            cmb3.AddItem "Charge Sales"
            cmb3.AddItem "Voucher Sales"
            cmb3.AddItem "Loyalty Sales"
            cmb3.AddItem "No Sales"
            cmb3.Text = "<All Transactions>"
            grdMain.Tag = ""
            Selection_Change
        Case "Purchase Journal"
            picMain.Visible = True
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            grdMain.Cols = 9
            grdMain.ColHidden(8) = True
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Dated"
            grdMain.TextMatrix(0, 1) = "User"
            grdMain.TextMatrix(0, 2) = "Supplier"
            grdMain.TextMatrix(0, 3) = "Document"
            grdMain.TextMatrix(0, 4) = "Transaction Type"
            grdMain.TextMatrix(0, 5) = "Total (excl) "
            grdMain.TextMatrix(0, 6) = "Tax "
            grdMain.TextMatrix(0, 7) = "Total (incl) "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignLeftCenter
            grdMain.ColAlignment(3) = flexAlignLeftCenter
            grdMain.ColAlignment(4) = flexAlignLeftCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.08
            grdMain.ColWidth(1) = grdMain.Width * 0.12
            grdMain.ColWidth(2) = grdMain.Width * 0.2
            grdMain.ColWidth(3) = grdMain.Width * 0.08
            grdMain.ColWidth(4) = grdMain.Width * 0.22
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdMain.ColWidth(7) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            cmb1.Clear
            lblCaption.Caption = "Reports - Purchase Journal"
            ActiveReadServer "Select Supplier_No,Supplier_Name from suppliers order by Supplier_name"
            cmb1.AddItem "<All Suppliers>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Suppliers>"
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            grdMain.Tag = ""
            Selection_Change
        Case "Sales Analysis (Graph)"
            picAnalysis.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            lblCaption.Caption = "Reports - Sales Analysis"
            cmb1.Clear
            cmb1.AddItem "Daily"
            cmb1.AddItem "Monthly"
            cmb1.Text = "Daily"
        Case "Sales Analysis by Product", "Sales Analysis by Department", "Sales Analysis by User", "Sales Analysis by Location", "Sales Analysis by Supplier", "Sales Analysis by Debtor"
            frmMain.Toolbar1.Buttons(14).Enabled = True
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
            cmb1.Left = 10680
            cmb3.Visible = False
            lblCaption.Caption = "Reports - " & cmb2.Text
            grdMain.Cols = 8
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " Product Code"
            grdMain.TextMatrix(0, 1) = "Description"
            grdMain.TextMatrix(0, 2) = "Qty Sold "
            grdMain.TextMatrix(0, 3) = "Total Cost "
            grdMain.TextMatrix(0, 4) = "GP% "
            grdMain.TextMatrix(0, 5) = "GP Value "
            grdMain.TextMatrix(0, 6) = "Total Revenue "
            grdMain.TextMatrix(0, 7) = "Profit Index "
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignRightCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.25
            grdMain.ColWidth(2) = grdMain.Width * 0.1
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.1
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdMain.ColWidth(7) = grdMain.Width * 0.1
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
            grdMain.Rows = 1
            grdMain.Tag = "1"
            Select Case cmb2.Text
                Case "Sales Analysis by Debtor"
                    cmb1.Clear
                    ActiveReadServer "Select Debtor_No,Debtor_Name from Debtors order by Debtor_name"
                    cmb1.AddItem "<All Debtors>"
                    While Not rs.EOF
                        cmb1.AddItem rs.Fields("Debtor_No") & " - " & rs.Fields("Debtor_Name")
                        rs.MoveNext
                    Wend
                    rs.Close
                    cmb1.Text = "<All Debtors>"
                Case "Sales Analysis by Product"
                    cmb1.Clear
                    cmb1.AddItem "<All Products>"
                    cmb1.AddItem "Star"
                    cmb1.AddItem "Cash Cow"
                    cmb1.AddItem "Dog"
                    cmb1.AddItem "Problem Child"
                    cmb1.Text = "<All Products>"
                Case "Sales Analysis by Department"
                    cmb1.Clear
                    ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
                    cmb1.AddItem "<All Departments>"
                    While Not rs.EOF
                        cmb1.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                        rs.MoveNext
                    Wend
                    rs.Close
                    cmb1.Text = "<All Departments>"
                Case "Sales Analysis by User"
                    cmb1.Width = 1845
                    cmb2.Width = 2415
                    cmb3.Width = 1905
                    cmb1.Left = 9840
                    cmb2.Left = 7380
                    cmb3.Left = 11730
                    cmb3.Visible = True
                    cmb1.Clear
                    ActiveReadServer "Select user_No,User_Name from Users order by user_no"
                    cmb1.AddItem "<All Users>"
                    While Not rs.EOF
                        cmb1.AddItem rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                        rs.MoveNext
                    Wend
                    cmb1.Text = "<All Users>"
                    rs.Close
                    cmb3.Clear
                    ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
                    cmb3.AddItem "<All Departments>"
                    While Not rs.EOF
                        cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                        rs.MoveNext
                    Wend
                    rs.Close
                    cmb3.Text = "<All Departments>"
                Case "Sales Analysis by Location"
                    cmb1.Clear
                    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
                    cmb1.AddItem "<All Locations>"
                    While Not rs.EOF
                        cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                        rs.MoveNext
                    Wend
                    rs.Close
                    cmb1.Text = "<All Locations>"
                Case "Sales Analysis by Supplier"
                    cmb1.Clear
                    ActiveReadServer "Select Supplier_No,Supplier_Name from Suppliers order by Supplier_name"
                    cmb1.AddItem "<All Suppliers>"
                    While Not rs.EOF
                        cmb1.AddItem rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
                        rs.MoveNext
                    Wend
                    rs.Close
                    cmb1.Text = "<All Suppliers>"
            End Select
            grdMain.Tag = ""
            Selection_Change
        Case "Trade Analysis"
            grdMain.Tag = "1"
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Locations>"
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            picTrade.Visible = True
            grdRev.ColAlignment(0) = flexAlignLeftCenter
            lblCaption.Caption = "Reports - Trade Analysis"
            grdRev.TextMatrix(0, 0) = "Revenue"
            grdRev.TextMatrix(0, 1) = "Revenue"
            grdRev.TextMatrix(0, 2) = "Revenue"
            grdRev.MergeRow(0) = True
            grdRev.Select 0, 0, 0, 1
            grdRev.Select 0, 0, 0, 2
            grdRev.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdRev.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            grdRev.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
            grdRev.TextMatrix(1, 0) = "Cash Sales:"
            grdRev.TextMatrix(2, 0) = "Card Sales:"
            grdRev.TextMatrix(3, 0) = "Voucher Sales:"
            grdRev.TextMatrix(4, 0) = "Charge Sales:"
            grdRev.TextMatrix(5, 0) = "Loyalty Sales:"
            grdRev.TextMatrix(6, 0) = "Revenue Total:"
            grdRev.ColAlignment(1) = flexAlignRightCenter
            grdRev.Cell(flexcpFontBold, 6, 0, 6, 1) = True
            grdRev.TextMatrix(1, 1) = "0.00"
            grdRev.TextMatrix(2, 1) = "0.00"
            grdRev.TextMatrix(3, 1) = "0.00"
            grdRev.TextMatrix(4, 1) = "0.00"
            grdRev.TextMatrix(5, 1) = "0.00"
            grdRev.TextMatrix(6, 1) = "0.00"
            grdRev.TextMatrix(1, 2) = "0"
            grdRev.TextMatrix(2, 2) = "0"
            grdRev.TextMatrix(3, 2) = "0"
            grdRev.TextMatrix(4, 2) = "0"
            grdRev.TextMatrix(5, 2) = "0"
            grdRev.TextMatrix(6, 2) = "0"
            grdRev.Select 6, 1, 6, 2
            grdRev.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdRev.Cell(flexcpBackColor, 6, 1, 6, 2) = &HE0E0E0
            grdTrans.TextMatrix(0, 0) = "Revenue Transactions"
            grdTrans.TextMatrix(0, 1) = "Revenue Transactions"
            grdTrans.TextMatrix(0, 2) = "Revenue Transactions"
            grdTrans.Select 0, 0, 0, 2
            grdTrans.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdTrans.MergeRow(0) = True
            grdTrans.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            grdTrans.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
            grdTrans.TextMatrix(1, 0) = "Payouts:"
            grdTrans.TextMatrix(2, 0) = "Deposits:"
            grdTrans.TextMatrix(3, 0) = "Recieve on Account:"
            grdTrans.TextMatrix(4, 0) = "Sub Total:"
            grdTrans.TextMatrix(5, 0) = "Total Reported:"
            grdTrans.ColAlignment(1) = flexAlignRightCenter
            grdTrans.Cell(flexcpFontBold, 4, 0, 4, 1) = True
            grdTrans.Cell(flexcpFontBold, 5, 0, 5, 1) = True
            grdTrans.TextMatrix(1, 1) = "0.00"
            grdTrans.TextMatrix(2, 1) = "0.00"
            grdTrans.TextMatrix(3, 1) = "0.00"
            grdTrans.TextMatrix(4, 1) = "0.00"
            grdTrans.TextMatrix(5, 1) = "0.00"
            grdTrans.TextMatrix(1, 2) = "0"
            grdTrans.TextMatrix(2, 2) = "0"
            grdTrans.TextMatrix(3, 2) = "0"
            grdTrans.TextMatrix(4, 2) = "0"
            grdTrans.TextMatrix(5, 2) = "0"
            grdTrans.Cell(flexcpBackColor, 4, 1, 4, 2) = &HE0E0E0
            grdTrans.Select 4, 1, 4, 2
            grdTrans.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdTrans.Select 5, 1, 5, 2
            grdTrans.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdTrans.Cell(flexcpBackColor, 5, 1, 5, 2) = &HE0E0E0
            
            grdCount.TextMatrix(0, 0) = "Other Transactions"
            grdCount.TextMatrix(0, 1) = "Other Transactions"
            grdCount.TextMatrix(0, 2) = "Other Transactions"
            grdCount.Select 0, 0, 0, 2
            grdCount.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdCount.MergeRow(0) = True
            grdCount.Cell(flexcpFontBold, 0, 0, 0, 2) = True
            grdCount.Cell(flexcpBackColor, 0, 0, 0, 2) = &HE0E0E0
            grdCount.TextMatrix(1, 0) = "Item Corrects:"
            grdCount.TextMatrix(2, 0) = "Voids:"
            grdCount.TextMatrix(3, 0) = "Returns:"
            grdCount.TextMatrix(4, 0) = "Wastages:"
            grdCount.TextMatrix(5, 0) = "Discount%:"
            grdCount.TextMatrix(6, 0) = "Discount Value":
            grdCount.TextMatrix(7, 0) = "Service Charges:"
            grdCount.ColAlignment(1) = flexAlignRightCenter
            grdCount.TextMatrix(1, 1) = "0.00"
            grdCount.TextMatrix(2, 1) = "0.00"
            grdCount.TextMatrix(3, 1) = "0.00"
            grdCount.TextMatrix(4, 1) = "0.00"
            grdCount.TextMatrix(5, 1) = "0.00"
            grdCount.TextMatrix(6, 1) = "0.00"
            grdCount.TextMatrix(7, 1) = "0.00"
            grdCount.TextMatrix(1, 2) = "0"
            grdCount.TextMatrix(2, 2) = "0"
            grdCount.TextMatrix(3, 2) = "0"
            grdCount.TextMatrix(4, 2) = "0"
            grdCount.TextMatrix(5, 2) = "0"
            grdCount.TextMatrix(6, 2) = "0"
            grdCount.TextMatrix(7, 2) = "0"
            
            grdTax.TextMatrix(0, 0) = "Tax Counters"
            grdTax.TextMatrix(0, 1) = "Tax Counters"
            grdTax.Select 0, 0, 0, 1
            grdTax.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdTax.MergeRow(0) = True
            grdTax.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            grdTax.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
            grdTax.TextMatrix(1, 0) = "Taxable Sales:"
            grdTax.TextMatrix(2, 0) = "Non-Taxable Sales:"
            grdTax.TextMatrix(3, 0) = "Tax Collected:"
            grdTax.TextMatrix(4, 0) = "Tax Calculated:"
            grdTax.TextMatrix(5, 0) = "Purchase Tax:"
            grdTax.ColAlignment(1) = flexAlignRightCenter
            grdTax.TextMatrix(1, 1) = "0.00"
            grdTax.TextMatrix(2, 1) = "0.00"
            grdTax.TextMatrix(3, 1) = "0.00"
            grdTax.TextMatrix(4, 1) = "0.00"
            grdTax.TextMatrix(5, 1) = "0.00"
            
            grdStock.TextMatrix(0, 0) = "Stock"
            grdStock.TextMatrix(0, 1) = "Stock"
            grdStock.Select 0, 0, 0, 1
            grdStock.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdStock.MergeRow(0) = True
            grdStock.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            grdStock.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
            grdStock.TextMatrix(1, 0) = "Opening Stock:"
            grdStock.TextMatrix(2, 0) = "-Consumption & Stock Variances:"
            grdStock.TextMatrix(3, 0) = "-Transfer Out:"
            grdStock.TextMatrix(4, 0) = "+Transfers In:"
            grdStock.TextMatrix(5, 0) = "+Purchases:"
            grdStock.TextMatrix(6, 0) = "Closing Stock:"
            grdStock.ColAlignment(1) = flexAlignRightCenter
            grdStock.TextMatrix(1, 1) = "0.00"
            grdStock.TextMatrix(2, 1) = "0.00"
            grdStock.TextMatrix(3, 1) = "0.00"
            grdStock.TextMatrix(4, 1) = "0.00"
            grdStock.TextMatrix(5, 1) = "0.00"
            grdStock.TextMatrix(6, 1) = "0.00"
            grdStock.Select 6, 1, 6, 1
            grdStock.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdStock.Cell(flexcpBackColor, 6, 1, 6, 1) = &HE0E0E0
            grdStock.Cell(flexcpFontBold, 6, 0, 6, 1) = True
            
            grdCred.TextMatrix(0, 0) = "Debtors"
            grdCred.TextMatrix(0, 1) = "Debtors"
            grdCred.TextMatrix(0, 2) = "Creditors"
            grdCred.TextMatrix(0, 3) = "Creditors"
            grdCred.Select 0, 0, 0, 3
            grdCred.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdCred.MergeRow(0) = True
            grdCred.Cell(flexcpFontBold, 0, 0, 0, 3) = True
            grdCred.Cell(flexcpBackColor, 0, 0, 0, 3) = &HE0E0E0
            grdCred.TextMatrix(1, 0) = "Opening Balance:"
            grdCred.TextMatrix(2, 0) = "+Sales:"
            grdCred.TextMatrix(3, 0) = "-Journals:"
            grdCred.TextMatrix(4, 0) = "-Receive on Account:"
            grdCred.TextMatrix(5, 0) = "Closing Balance:"
            grdCred.TextMatrix(1, 2) = "Opening Balance:"
            grdCred.TextMatrix(2, 2) = "+Purchases:"
            grdCred.TextMatrix(3, 2) = "-Journals:"
            grdCred.TextMatrix(4, 2) = "-Payments:"
            grdCred.TextMatrix(5, 2) = "Closing Balance:"
            grdCred.ColAlignment(1) = flexAlignRightCenter
            grdCred.TextMatrix(1, 1) = "0.00"
            grdCred.TextMatrix(2, 1) = "0.00"
           ' grdCred.TextMatrix(3, 1) = "0.00"
            grdCred.TextMatrix(4, 1) = "0.00"
            grdCred.TextMatrix(5, 1) = "0.00"
            grdCred.TextMatrix(1, 3) = "0.00"
            grdCred.TextMatrix(2, 3) = "0.00"
            grdCred.TextMatrix(3, 3) = "0.00"
            grdCred.TextMatrix(4, 3) = "0.00"
            grdCred.TextMatrix(5, 3) = "0.00"
            grdCred.Select 5, 1, 5, 1
            grdCred.CellBorder &H808080, 1, 1, 1, 0, 1, 1
            grdCred.Select 5, 3, 5, 3
            grdCred.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdCred.Cell(flexcpBackColor, 5, 1, 5, 1) = &HE0E0E0
            grdCred.Cell(flexcpBackColor, 5, 3, 5, 3) = &HE0E0E0
            grdCred.Cell(flexcpBackColor, 1, 1, 4, 1) = &HE9E9E9
            grdCred.Cell(flexcpBackColor, 1, 3, 4, 3) = &HE9E9E9
            grdCred.Cell(flexcpFontBold, 5, 0, 5, 3) = True
            grdCred.ColAlignment(0) = flexAlignRightCenter
            grdCred.ColAlignment(2) = flexAlignRightCenter
            
            grdGP.Select 0, 0, 1, 0
            grdGP.CellBorder &H808080, 0, 0, 1, 0, 1, 0
            grdGP.MergeRow(0) = True
            grdGP.Cell(flexcpFontBold, 0, 0, 1, 3) = True
            grdGP.Cell(flexcpBackColor, 0, 1, 1, 1) = &HE0E0E0
            grdGP.Cell(flexcpBackColor, 0, 3, 1, 3) = &HE0E0E0
            grdGP.TextMatrix(0, 0) = "GP Percentage (%):"
            grdGP.TextMatrix(1, 0) = "GP Value:"
            grdGP.TextMatrix(0, 2) = "Customer Count:"
            grdGP.TextMatrix(1, 2) = "Spend per Head:"
            grdGP.ColAlignment(1) = flexAlignRightCenter
            grdGP.TextMatrix(0, 3) = "0"
            grdGP.TextMatrix(1, 3) = "0.00"
            grdGP.TextMatrix(0, 1) = "0%"
            grdGP.TextMatrix(1, 1) = "0.00"
            grdGP.ColAlignment(0) = flexAlignRightCenter
            grdGP.ColAlignment(2) = flexAlignRightCenter
            grdRev.SetFocus
            grdMain.Tag = ""
            Selection_Change
        Case "Sales Analysis by Department"
            lblCaption.Caption = "Reports - Sales Analysis by Department"
        Case "Sales Analysis by Location"
            lblCaption.Caption = "Reports - Sales Analysis by Location"
        Case "Sales Analysis by User"
            lblCaption.Caption = "Reports - Sales Analysis by User"
        Case "Daily Trade Analysis"
            daytradechange

        Case "Sales Analysis by User type"
            Screen.MousePointer = 11
            lblCaption.Caption = "Reports - Sales Analysis by User type"
                cmb3.Left = 51730
                cmb1.Left = 51730
           
              Timer1.Enabled = True
            'Selection_Change
           
         Case "Trade Comparison" ' Trade Comparison for Beach Hotel
             
'             Pictime.Left = 9835
'             Pictime.top = 660
            
             cmb2.Width = 2415
           
             Lbltimeback.Left = 9860
             Lbltimeback.top = 660
            Lbltime.Left = 9980
            Lbltime.top = 730
            
             Lbltime.Caption = Format(Time_Start, "Medium Time") & " to " & Format(Time_Stop, "Medium Time")
            TStart.Value = Time_Start
            TEnd.Value = Time_Stop
            Buttime.Visible = True
            Lbltime.Visible = True
            Lbltimeback.Visible = True
             TradeComparison
                
            
    End Select
    Screen.MousePointer = 0
    On Error GoTo 0
End Sub






Private Sub Sabutreportgetready() ' Sales by user type report
'*********************************************************************

            
            
            
            
            
            If Screen.MousePointer <> 11 Then Screen.MousePointer = 11
            cmb2.Text = "Sales Analysis by User type"
            frmMain.Toolbar1.Buttons(14).Enabled = True
            picMain.Visible = True
            cmb2.Width = 3225
            cmb1.Width = 2925
            cmb2.Left = 7380
     
            cmb3.Visible = False
            lblCaption.Caption = "Reports - " & cmb2.Text
            grdMain.Cols = 9
            grdMain.FixedCols = 0
            grdMain.TextMatrix(0, 0) = " User No"
            grdMain.TextMatrix(0, 1) = " User Name"
            grdMain.TextMatrix(0, 2) = " Tendercount"
            grdMain.TextMatrix(0, 3) = " -Card"
            grdMain.TextMatrix(0, 4) = " -Voucher"
            grdMain.TextMatrix(0, 5) = " -Charge"
            grdMain.TextMatrix(0, 6) = " Cash"
            grdMain.TextMatrix(0, 7) = " Total"
            grdMain.ColAlignment(0) = flexAlignLeftCenter
            grdMain.ColAlignment(1) = flexAlignLeftCenter
            grdMain.ColAlignment(2) = flexAlignRightCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            grdMain.ColAlignment(5) = flexAlignRightCenter
            grdMain.ColAlignment(6) = flexAlignRightCenter
            grdMain.ColAlignment(7) = flexAlignRightCenter
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.25
            grdMain.ColWidth(2) = grdMain.Width * 0.1
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.1
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColWidth(6) = grdMain.Width * 0.1
            grdMain.ColWidth(7) = grdMain.Width * 0.1
            grdMain.ColHidden(8) = True
            grdTotal.Cols = grdMain.Cols
            For i = 0 To grdMain.Cols - 1
                grdTotal.ColAlignment(i) = grdMain.ColAlignment(i)
                grdTotal.ColWidth(i) = grdMain.ColWidth(i)
            Next i
             grdMain.Rows = 1
            grdMain.Tag = "1"
            
            
        
            
            If Screen.MousePointer <> 11 Then Screen.MousePointer = 11
 If Right(Str(Time_Stop), 2) = "AM" Then
                Selender = DateAdd("d", 1, mthViewEnd.Value)
                lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
            Else
                Selender = mthViewEnd.Value
            End If
           
            LocString = "%"
            DeptString = "%"
            grdMain.Rows = grdMain.Rows + 1
          
            grdMain.TextMatrix(1, 0) = "Usertype:"
            grdMain.TextMatrix(1, 1) = "Barman"
        
    '**************************************************************************
            'Remember to put in dates and time in where clause!!!!!!!!!!!!!!!!1
        

            ' For each user that worked
'            Case "Manager": Vtype = 0
'            Case "Night Manager": Vtype = 1
'            Case "Reservations Clerk": Vtype = 2
'            Case "Waiter": Vtype = 3
'            Case "Barman": Vtype = 4
'            Case "GRV Clerk": Vtype = 5
'            Case "Buyer": Vtype = 6
'            Case "Supervisor": Vtype = 7
'            Case "Cashier": Vtype = 8
'            Case "Owner": Vtype = 9
'            Case "Staff Member": Vtype = 10
                         
            
            
            ' Barman
            ActiveReadServer "SELECT Function_Key, CASE Function_Key WHEN '9' THEN 'Cash' WHEN '10' THEN 'Card' WHEN '11' THEN 'Voucher' WHEN '12' THEN 'Charge' WHEN '13' THEN 'Loyalty' END" & _
            " AS Sale_Type, COUNT(Line_No) AS Tend_Count, SUM(Sales_Tax) AS Sales_Tax, SUM(Line_Total) AS Line_Total, User_name,dbo.users.user_no, dbo.Sales_Journal_View.User_Type AS User_type " & _
            " From Sales_Journal_View Inner Join dbo.Users ON dbo.Sales_Journal_View.User_No = dbo.Users.User_No" & _
            " WHERE     (Function_Key IN (9, 10, 11, 12, 13)) AND dbo.Sales_Journal_View.User_type = '4' " & _
            " and (Date_Time BETWEEN '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and  '" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
            " GROUP BY Function_Key, dbo.Sales_Journal_View.User_No , dbo.Users.User_Name, dbo.users.user_no, dbo.Sales_Journal_View.User_Type " & _
            " ORDER BY dbo.Sales_Journal_View.User_No, dbo.Sales_Journal_View.User_Type"
            
               
             While Not rs.EOF
             grdMain.Rows = grdMain.Rows + 1

            If grdMain.TextMatrix((grdMain.Rows - 2), 0) = rs.Fields("user_no") Then grdMain.Rows = grdMain.Rows - 1
           Select Case rs.Fields("Sale_Type")
            Case "Cash"
                
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 6) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(grdMain.TextMatrix((grdMain.Rows - 1), 2)) + Val(rs.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs.Fields("User_Type")
            Case "Card"
                
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 3) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(grdMain.TextMatrix((grdMain.Rows - 1), 2)) + Val(rs.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs.Fields("User_Type")
            Case "Voucher"
                
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 4) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(grdMain.TextMatrix((grdMain.Rows - 1), 2)) + Val(rs.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs.Fields("User_Type")
            Case "Charge"
                
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 5) = Format(Val(rs.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(grdMain.TextMatrix((grdMain.Rows - 1), 2)) + Val(rs.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs.Fields("User_Type")
'            Case "Loyalty"
'                grdMain.TextMatrix(5, 1) = Format(Val(rs.Fields("Line_Total")), "0.00")
'                grdMain.TextMatrix(5, 2) = Val(rs.Fields("Tend_Count") & "")
        End Select
           
         grdMain.TextMatrix((grdMain.Rows - 1), 7) = Format((grdMain.ValueMatrix((grdMain.Rows - 1), 3) + grdMain.ValueMatrix((grdMain.Rows - 1), 4) + grdMain.ValueMatrix((grdMain.Rows - 1), 5) + grdMain.ValueMatrix((grdMain.Rows - 1), 6)), "00.00")
         rs.MoveNext
            Wend
            rs.Close
            



           
            
' 2nd part waiters

grdMain.Rows = grdMain.Rows + 1
grdMain.TextMatrix((grdMain.Rows - 1), 0) = "Usertype:"
grdMain.TextMatrix((grdMain.Rows - 1), 1) = "Waiters"

' Waiters


       ActiveReadServer2 "SELECT Function_Key, CASE Function_Key WHEN '9' THEN 'Cash' WHEN '10' THEN 'Card' WHEN '11' THEN 'Voucher' WHEN '12' THEN 'Charge' WHEN '13' THEN 'Loyalty' END" & _
            " AS Sale_Type, COUNT(Line_No) AS Tend_Count, SUM(Sales_Tax) AS Sales_Tax, SUM(Line_Total) AS Line_Total, User_name,dbo.users.user_no, dbo.Sales_Journal_View.User_Type AS User_type " & _
            " From Sales_Journal_View Inner Join dbo.Users ON dbo.Sales_Journal_View.User_No = dbo.Users.User_No" & _
            " WHERE     (Function_Key IN (9, 10, 11, 12, 13)) AND dbo.Sales_Journal_View.User_type = '3' " & _
            " and (Date_Time BETWEEN '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and  '" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
            " GROUP BY Function_Key, dbo.Sales_Journal_View.User_No , dbo.Users.User_Name, dbo.users.user_no, dbo.Sales_Journal_View.User_Type " & _
            " ORDER BY dbo.Sales_Journal_View.User_No, dbo.Sales_Journal_View.User_Type"


             While Not rs2.EOF
             grdMain.Rows = grdMain.Rows + 1

        If grdMain.TextMatrix((grdMain.Rows - 2), 0) = rs2.Fields("user_no") Then grdMain.Rows = grdMain.Rows - 1
        Select Case rs2.Fields("Sale_Type")
            Case "Cash"
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs2.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs2.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 6) = Format(Val(rs2.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(rs2.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs2.Fields("User_Type")
            Case "Card"
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs2.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs2.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 3) = Format(Val(rs2.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(rs2.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs2.Fields("User_Type")
            Case "Voucher"
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs2.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs2.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 4) = Format(Val(rs2.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(rs2.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs2.Fields("User_Type")
            Case "Charge"
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs2.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs2.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 5) = Format(Val(rs2.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(rs2.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs2.Fields("User_Type")
        End Select
        grdMain.TextMatrix((grdMain.Rows - 1), 7) = Format((grdMain.ValueMatrix((grdMain.Rows - 1), 3) + grdMain.ValueMatrix((grdMain.Rows - 1), 4) + grdMain.ValueMatrix((grdMain.Rows - 1), 5) + grdMain.ValueMatrix((grdMain.Rows - 1), 6)), "00.00")
        rs2.MoveNext
        Wend
        rs2.Close
        
        'Managers
        grdMain.Rows = grdMain.Rows + 1
grdMain.TextMatrix((grdMain.Rows - 1), 0) = "Usertype:"
grdMain.TextMatrix((grdMain.Rows - 1), 1) = "Managers"

         ActiveReadServer3 "SELECT Function_Key, CASE Function_Key WHEN '9' THEN 'Cash' WHEN '10' THEN 'Card' WHEN '11' THEN 'Voucher' WHEN '12' THEN 'Charge' WHEN '13' THEN 'Loyalty' END" & _
            " AS Sale_Type, COUNT(Line_No) AS Tend_Count, SUM(Sales_Tax) AS Sales_Tax, SUM(Line_Total) AS Line_Total, User_name,dbo.users.user_no, dbo.Sales_Journal_View.User_Type AS User_type " & _
            " From Sales_Journal_View Inner Join dbo.Users ON dbo.Sales_Journal_View.User_No = dbo.Users.User_No" & _
            " WHERE     (Function_Key IN (9, 10, 11, 12, 13)) AND dbo.Sales_Journal_View.User_type = '0' " & _
            " and (Date_Time BETWEEN '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and  '" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
            " GROUP BY Function_Key, dbo.Sales_Journal_View.User_No , dbo.Users.User_Name, dbo.users.user_no, dbo.Sales_Journal_View.User_Type " & _
            " ORDER BY dbo.Sales_Journal_View.User_No, dbo.Sales_Journal_View.User_Type"
             While Not rs3.EOF
             grdMain.Rows = grdMain.Rows + 1

        If grdMain.TextMatrix((grdMain.Rows - 2), 0) = rs3.Fields("user_no") Then grdMain.Rows = grdMain.Rows - 1
        Select Case rs3.Fields("Sale_Type")
            Case "Cash"
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs3.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs3.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 6) = Format(Val(rs3.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(rs3.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs3.Fields("User_Type")
            Case "Card"
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs3.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs3.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 3) = Format(Val(rs3.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(rs3.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs3.Fields("User_Type")
            Case "Voucher"
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs3.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs3.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 4) = Format(Val(rs3.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(rs3.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs3.Fields("User_Type")
            Case "Charge"
                grdMain.TextMatrix((grdMain.Rows - 1), 0) = rs3.Fields("user_no")
                grdMain.TextMatrix((grdMain.Rows - 1), 1) = rs3.Fields("User_Name")
                grdMain.TextMatrix((grdMain.Rows - 1), 5) = Format(Val(rs3.Fields("Line_Total")), "0.00")
                grdMain.TextMatrix((grdMain.Rows - 1), 2) = Val(rs3.Fields("Tend_Count") & "")
                grdMain.TextMatrix((grdMain.Rows - 1), 8) = rs3.Fields("User_Type")
        End Select

           grdMain.TextMatrix((grdMain.Rows - 1), 7) = Format((grdMain.ValueMatrix((grdMain.Rows - 1), 3) + grdMain.ValueMatrix((grdMain.Rows - 1), 4) + grdMain.ValueMatrix((grdMain.Rows - 1), 5) + grdMain.ValueMatrix((grdMain.Rows - 1), 6)), "00.00")
        rs3.MoveNext
        Wend
        rs3.Close
        
        If grdMain.Rows > 1 Then
            grdMain.Row = 1
            grdMain.SetFocus
            End If
          ActiveUpdateServer "Delete  from Sabutemp"
        frmMain.Toolbar1.Buttons(16).Enabled = True
        frmMain.Toolbar1.Buttons(16).Tag = "Reports"
        grdMain.Tag = ""
            
       
       For f = 1 To grdMain.Rows
       If grdMain.TextMatrix((f), 0) <> " User No" Then
       If grdMain.TextMatrix((f), 0) <> " Usertype:" Then
       ActiveUpdateServer "Insert into Sabutemp (User_no,User_name,Cashed,Card,Cheque,Charge, User_Type, Total) values ('" & grdMain.TextMatrix((f), 0) & "', '" & grdMain.TextMatrix((f), 1) & "', '" & Format(Val(grdMain.TextMatrix((f), 6)), "0.00") & "', '" & Format(Val(grdMain.TextMatrix((f), 3)), "0.00") & "', '" & Format(Val(grdMain.TextMatrix((f), 4)), "0.00") & "', '" & Format(Val(grdMain.TextMatrix((f), 5)), "0.00") & "', '" & grdMain.TextMatrix((f), 8) & "', ' " & Format(Val(grdMain.TextMatrix((f), 7)), "0.00") & "' )"
       End If
       End If
       
       Next f

            
End Sub
' Compare last year's sales to current sales
Private Sub TradeComparison()
If Screen.MousePointer <> 11 Then Screen.MousePointer = 11
            cmb2.Text = "Trade Comparison"
            cmb1.Visible = False
            cmb3.Visible = False
            Buttime.Left = 12180
            Buttime.top = 660
            frmMain.Toolbar1.Buttons(14).Enabled = True
            Buttime.Visible = True
            Lbltime.Visible = True
            Lbltimeback.Visible = True
            PicTradecomparison.Visible = True
            Tradecomparison_Changed
          
End Sub
Private Sub Tradecomparison_Changed()
           Dim Counts As Integer, Breakfaststart As String, _
           Breakfastend As String, Lunchstart As String, Lunchend As String, Dinnerstart As String
           Dim Lastyear
           Dim B_total As Double
           Dim B_total_y As Double
           Dim Lunch_total As Double
           Dim Lunch_total_y As Double
           Dim Dinner_total As Double
           Dim Dinner_total_y As Double
           
           
            txtcount.BackColor = &HE0E0E0
            Breakfaststart = "03:59:59 AM" 'constant
            Breakfastend = "10:59:59 AM" 'constant
            Lunchstart = "11:00:00 AM" 'constant
            Lunchend = "03:59:59 PM" 'constant
            Dinnerstart = "04:00:00 PM" 'constant
            Lastyear = DateAdd("YYYY", -1, mthViewStart.Value)
            ' Total Sales count ((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
            ActiveReadServer2 "SELECT Function_Key," & _
            "Case Function_Key" & _
            " WHEN 9 THEN 'Cash'" & _
            " WHEN 10 THEN 'Card'" & _
            " WHEN 11 THEN 'Voucher'" & _
            " WHEN 12 THEN 'Charge'" & _
            " WHEN 13 THEN 'Loyalty'" & _
            " END As Sale_Type" & _
            ",COUNT(Line_No) AS Tend_Count,SUM(Sales_Tax) AS Sales_Tax,SUM(Line_Total) AS Line_Total,Sum(Ave_Cost) as Ave_Cost" & _
            " From Sales_Journal" & _
            " WHERE (Function_Key IN ( 9, 10, 11, 12, 13)) " & _
            " and (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')" & _
            " GROUP BY Function_Key"
            Counts = 0
            While Not rs2.EOF
            Select Case rs2.Fields("Sale_Type")
                Case "Cash"
                    Counts = Counts + Val(rs2.Fields("Tend_Count") & "")
                Case "Card"
                    Counts = Counts + Val(rs2.Fields("Tend_Count") & "")
                Case "Voucher"
                    Counts = Counts + Val(rs2.Fields("Tend_Count") & "")
                Case "Charge"
                   Counts = Counts + Val(rs2.Fields("Tend_Count") & "")
                Case "Loyalty"
                    Counts = Counts + Val(rs2.Fields("Tend_Count") & "")
            End Select
            rs2.MoveNext
                Wend
                rs2.Close
                txtcount.Text = "Total Sales count = " & Counts
            '(((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
           
' Grd Today 888888888888888888888888888888888888888888888888888
           lblCaption.Caption = "Reports - Trade Comparison"
           LblCaption2.Caption = "Trade Comparison For " & Format(Lastyear, "DD MMM YYYY") & " And " & Format(mthViewStart.Value, "DD MMM YYYY")
        'Gridtoday
            GrdToday.Cell(flexcpFontBold, 1, 0, 6, 0) = True
            GrdToday.Cell(flexcpFontBold, 1, 0, 0, 5) = True
            Dim i As Integer
                    For i = 0 To 5
                    GrdToday.ColWidth(i) = GrdToday.Width * 0.166
                    Next i
                GrdToday.TextMatrix(0, 0) = "Sales Breakdown"
                GrdToday.TextMatrix(0, 1) = "Sales"
                GrdToday.TextMatrix(0, 2) = "Sales History"
                GrdToday.TextMatrix(0, 3) = "Month To Date"
                GrdToday.TextMatrix(0, 4) = " "
                GrdToday.TextMatrix(0, 5) = "Variance"
                GrdToday.TextMatrix(1, 0) = "Breakfast"
                GrdToday.TextMatrix(2, 0) = "Lunch"
                GrdToday.TextMatrix(3, 0) = "Dinner"
                GrdToday.TextMatrix(4, 0) = "SUB TOTAL FOOD"
                GrdToday.TextMatrix(5, 0) = "Bar"
                GrdToday.TextMatrix(6, 0) = "SUB TOTAL BAR"
                GrdToday.Cell(flexcpFontBold, 0, 0, 0, 1) = True
                GrdToday.Cell(flexcpBackColor, 0, 0, 0, 5) = &HE0E0E0
                GrdToday.Cell(flexcpBackColor, 0, 0, 6, 0) = &HE0E0E0
                
          ' Get Breakfasttotal
            ActiveReadServer1 "Select isnull(SUM(Line_Total),0) AS Breakfasttotal" & _
            " From Sales_Journal Where (Function_Key in(9,10,11,12,13))" & _
            " and (Date_Time > '" & mthViewStart.Value & " " & Breakfaststart & "' and Date_Time<'" & mthViewStart & " " & Breakfastend & "')"
            B_total = rs1.Fields("Breakfasttotal")
            rs1.Close
            GrdToday.TextMatrix(1, 1) = Format(B_total, "00.00")
            
            ' Get Breakfasttotallastyear
            ActiveReadServer1 "Select isnull(SUM(Line_Total),0) AS BreakfasttotalB" & _
            " From Sales_Journal Where (Function_Key in(9,10,11,12,13))" & _
            " and (Date_Time > '" & Lastyear & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Lastyear & " " & Breakfastend & "')"
            B_total_y = rs1.Fields("BreakfasttotalB")
            rs1.Close
            GrdToday.TextMatrix(1, 2) = Format(B_total_y, "00.00")
            
            
            ' Get Lunchtotal
            ActiveReadServer1 "Select isnull(SUM(Line_Total),0) AS Lunchtotal" & _
            " From Sales_Journal Where (Function_Key in(9,10,11,12,13))" & _
            " and (Date_Time > '" & mthViewStart.Value & " " & Lunchstart & "' and Date_Time<'" & mthViewStart & " " & Lunchend & "')"
            Lunch_total = rs1.Fields("Lunchtotal")
            rs1.Close
            GrdToday.TextMatrix(2, 1) = Format(Lunch_total, "00.00")
            
            ' Get Lunchtotallastyear
            ActiveReadServer1 "Select isnull(SUM(Line_Total),0) AS LunchtotalB" & _
            " From Sales_Journal Where (Function_Key in(9,10,11,12,13))" & _
            " and (Date_Time > '" & Lastyear & " " & Lunchstart & "' and Date_Time<'" & Lastyear & " " & Lunchend & "')"
            Lunch_total_y = rs1.Fields("LunchtotalB")
            rs1.Close
            GrdToday.TextMatrix(2, 2) = Format(Lunch_total_y, "00.00")
            ' Get Dinnertotal
            ActiveReadServer1 "Select isnull(SUM(Line_Total),0) AS Dinnertotal" & _
            " From Sales_Journal Where (Function_Key in(9,10,11,12,13))" & _
            " and (Date_Time > '" & mthViewStart.Value & " " & Dinnerstart & "' and Date_Time<'" & mthViewStart + 1 & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
            Dinner_total = rs1.Fields("Dinnertotal")
            rs1.Close
            GrdToday.TextMatrix(3, 1) = Format(Dinner_total, "00.00")
            ' Get Dinnertotallastyear
            ActiveReadServer1 "Select isnull(SUM(Line_Total),0) AS DinnertotalB" & _
            " From Sales_Journal Where (Function_Key in(9,10,11,12,13))" & _
            " and (Date_Time > '" & Lastyear & " " & Dinnerstart & "' and Date_Time<'" & Lastyear & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
            Dinner_total_y = rs1.Fields("DinnertotalB")
            rs1.Close
            GrdToday.TextMatrix(3, 2) = Format(Dinner_total_y, "00.00")
            
'8888888888888888888888888888888888888888888888888888888888888888
            
             GrdToday.TextMatrix(4, 1) = Format(GrdToday.ValueMatrix(1, 1) + GrdToday.ValueMatrix(2, 1) + GrdToday.ValueMatrix(3, 1), "00.00")
             GrdToday.TextMatrix(4, 2) = Format(GrdToday.ValueMatrix(1, 2) + GrdToday.ValueMatrix(2, 2) + GrdToday.ValueMatrix(3, 2), "00.00")
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            'GrdTodayTotal777777777777777777777777777777777777777777777777
            
            For i = 0 To 5
            GrdTodayTotal.ColWidth(i) = GrdTodayTotal.Width * 0.166
            Next i
            GrdTodayTotal.TextMatrix(0, 0) = "TOTAL SALES"
            GrdTodayTotal.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            GrdTodayTotal.Cell(flexcpBackColor, 0, 0, 0, 5) = &HE0E0E0
            GrdTodayTotal.Cell(flexcpBackColor, 0, 0, 1, 0) = &HE0E0E0
            '7777777777777777777777777777777777777777777777777777777777777
            
            '6666666666666666666666666666666666666666666666666666666666666
            'Grdcovers
            Dim f As Integer
            Grdcovers.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            Grdcovers.Cell(flexcpFontBold, 1, 0, 6, 0) = True
            For f = 0 To 6
            Grdcovers.ColWidth(f) = Grdcovers.Width * 0.166
            Next f
            Grdcovers.TextMatrix(0, 0) = "COVERS"
            Grdcovers.TextMatrix(2, 0) = "Breakfast"
            Grdcovers.TextMatrix(3, 0) = "Lunch"
            Grdcovers.TextMatrix(4, 0) = "Dinner"
            Grdcovers.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            Grdcovers.Cell(flexcpBackColor, 0, 0, 0, 5) = &HE0E0E0
            Grdcovers.Cell(flexcpBackColor, 0, 0, 6, 0) = &HE0E0E0
            '6666666666666666666666666666666666666666666666666666666666666
            
            
            '555555555555555555555555555555555555555555555555555555555555555
            'GrdAvespend
            For f = 0 To 6
            GrdAvespend.ColWidth(f) = GrdAvespend.Width * 0.166
            Next f
            GrdAvespend.TextMatrix(0, 0) = "AVERAGE SPEND"
            GrdAvespend.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            GrdAvespend.Cell(flexcpFontBold, 1, 0, 6, 0) = True
            GrdAvespend.Cell(flexcpBackColor, 0, 0, 0, 2) = &HE0E0E0
            GrdAvespend.Cell(flexcpBackColor, 0, 0, 0, 5) = &HE0E0E0
            GrdAvespend.Cell(flexcpBackColor, 0, 0, 6, 0) = &HE0E0E0
            '555555555555555555555555555555555555555555555555555555555555555
            
            
            '999999999999999999999999999999999999999999999999999999999999999
            'GrdCompare
            For f = 0 To 6
            GrdCompare.ColWidth(f) = GrdCompare.Width * 0.166
            Next f
            GrdCompare.TextMatrix(0, 0) = "ACTUAL vs BUDGET"
            GrdCompare.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            GrdCompare.Cell(flexcpFontBold, 1, 0, 6, 0) = True
            GrdCompare.Cell(flexcpBackColor, 0, 0, 0, 5) = &HE0E0E0
            GrdCompare.Cell(flexcpBackColor, 0, 0, 6, 0) = &HE0E0E0
           '99999999999999999999999999999999999999999999999999999999999999999
           
            
            
            PicTradecomparison.SetFocus
            
            
End Sub
Private Sub daytradechange()
grdMain.Tag = ""
            cmb1.Clear
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            cmb1.AddItem "<All Locations>"
            While Not rs.EOF
                cmb1.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb1.Text = "<All Locations>"
            cmb3.Clear
            ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
            cmb3.AddItem "<All Departments>"
            While Not rs.EOF
                cmb3.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
                rs.MoveNext
            Wend
            rs.Close
            cmb3.Text = "<All Departments>"
            cmb1.Width = 1845
            cmb2.Width = 2415
            cmb3.Width = 1905
            cmb1.Left = 9840
            cmb2.Left = 7380
            cmb3.Left = 11730
            cmb3.Visible = True
            picTrade.Visible = True
            grdRev.ColAlignment(0) = flexAlignLeftCenter
            lblCaption.Caption = "Reports - Daily Trade Analysis"
            grdRev.TextMatrix(0, 0) = "Revenue"
            grdRev.TextMatrix(0, 1) = "Revenue"
            grdRev.TextMatrix(0, 2) = "Revenue"
            grdRev.MergeRow(0) = True
            grdRev.Select 0, 0, 0, 1
            grdRev.Select 0, 0, 0, 2
            grdRev.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdRev.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            grdRev.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
            grdRev.TextMatrix(1, 0) = "Cash Sales:"
            grdRev.TextMatrix(2, 0) = "Card Sales:"
            grdRev.TextMatrix(3, 0) = "Voucher Sales:"
            grdRev.TextMatrix(4, 0) = "Charge Sales:"
            grdRev.TextMatrix(5, 0) = "Loyalty Sales:"
            grdRev.TextMatrix(6, 0) = "Revenue Total:"
            grdRev.ColAlignment(1) = flexAlignRightCenter
            grdRev.Cell(flexcpFontBold, 6, 0, 6, 1) = True
            grdRev.TextMatrix(1, 1) = "0.00"
            grdRev.TextMatrix(2, 1) = "0.00"
            grdRev.TextMatrix(3, 1) = "0.00"
            grdRev.TextMatrix(4, 1) = "0.00"
            grdRev.TextMatrix(5, 1) = "0.00"
            grdRev.TextMatrix(6, 1) = "0.00"
            grdRev.TextMatrix(1, 2) = "0"
            grdRev.TextMatrix(2, 2) = "0"
            grdRev.TextMatrix(3, 2) = "0"
            grdRev.TextMatrix(4, 2) = "0"
            grdRev.TextMatrix(5, 2) = "0"
            grdRev.TextMatrix(6, 2) = "0"
            grdRev.Select 6, 1, 6, 2
            grdRev.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdRev.Cell(flexcpBackColor, 6, 1, 6, 2) = &HE0E0E0
            grdTrans.TextMatrix(0, 0) = "Revenue Transactions"
            grdTrans.TextMatrix(0, 1) = "Revenue Transactions"
            grdTrans.TextMatrix(0, 2) = "Revenue Transactions"
            grdTrans.Select 0, 0, 0, 2
            grdTrans.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdTrans.MergeRow(0) = True
            grdTrans.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            grdTrans.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
            grdTrans.TextMatrix(1, 0) = "Payouts:"
            grdTrans.TextMatrix(2, 0) = "Deposits:"
            grdTrans.TextMatrix(3, 0) = "Recieve on Account:"
            grdTrans.TextMatrix(4, 0) = "Sub Total:"
            grdTrans.TextMatrix(5, 0) = "Total Reported:"
            grdTrans.ColAlignment(1) = flexAlignRightCenter
            grdTrans.Cell(flexcpFontBold, 4, 0, 4, 1) = True
            grdTrans.Cell(flexcpFontBold, 5, 0, 5, 1) = True
            grdTrans.TextMatrix(1, 1) = "0.00"
            grdTrans.TextMatrix(2, 1) = "0.00"
            grdTrans.TextMatrix(3, 1) = "0.00"
            grdTrans.TextMatrix(4, 1) = "0.00"
            grdTrans.TextMatrix(5, 1) = "0.00"
            grdTrans.TextMatrix(1, 2) = "0"
            grdTrans.TextMatrix(2, 2) = "0"
            grdTrans.TextMatrix(3, 2) = "0"
            grdTrans.TextMatrix(4, 2) = "0"
            grdTrans.TextMatrix(5, 2) = "0"
            grdTrans.Cell(flexcpBackColor, 4, 1, 4, 2) = &HE0E0E0
            grdTrans.Select 4, 1, 4, 2
            grdTrans.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdTrans.Select 5, 1, 5, 2
            grdTrans.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdTrans.Cell(flexcpBackColor, 5, 1, 5, 2) = &HE0E0E0
            
            grdCount.TextMatrix(0, 0) = "Other Transactions"
            grdCount.TextMatrix(0, 1) = "Other Transactions"
            grdCount.TextMatrix(0, 2) = "Other Transactions"
            grdCount.Select 0, 0, 0, 2
            grdCount.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdCount.MergeRow(0) = True
            grdCount.Cell(flexcpFontBold, 0, 0, 0, 2) = True
            grdCount.Cell(flexcpBackColor, 0, 0, 0, 2) = &HE0E0E0
            grdCount.TextMatrix(1, 0) = "Item Corrects:"
            grdCount.TextMatrix(2, 0) = "Voids:"
            grdCount.TextMatrix(3, 0) = "Returns:"
            grdCount.TextMatrix(4, 0) = "Wastages:"
            grdCount.TextMatrix(5, 0) = "Discount%:"
            grdCount.TextMatrix(6, 0) = "Discount Value":
            grdCount.TextMatrix(7, 0) = "Service Charges:"
            grdCount.ColAlignment(1) = flexAlignRightCenter
            grdCount.TextMatrix(1, 1) = "0.00"
            grdCount.TextMatrix(2, 1) = "0.00"
            grdCount.TextMatrix(3, 1) = "0.00"
            grdCount.TextMatrix(4, 1) = "0.00"
            grdCount.TextMatrix(5, 1) = "0.00"
            grdCount.TextMatrix(6, 1) = "0.00"
            grdCount.TextMatrix(7, 1) = "0.00"
            grdCount.TextMatrix(1, 2) = "0"
            grdCount.TextMatrix(2, 2) = "0"
            grdCount.TextMatrix(3, 2) = "0"
            grdCount.TextMatrix(4, 2) = "0"
            grdCount.TextMatrix(5, 2) = "0"
            grdCount.TextMatrix(6, 2) = "0"
            grdCount.TextMatrix(7, 2) = "0"
            
            grdTax.TextMatrix(0, 0) = "Tax Counters"
            grdTax.TextMatrix(0, 1) = "Tax Counters"
            grdTax.Select 0, 0, 0, 1
            grdTax.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdTax.MergeRow(0) = True
            grdTax.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            grdTax.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
            grdTax.TextMatrix(1, 0) = "Taxable Sales:"
            grdTax.TextMatrix(2, 0) = "Non-Taxable Sales:"
            grdTax.TextMatrix(3, 0) = "Tax Collected:"
            grdTax.TextMatrix(4, 0) = "Tax Calculated:"
            grdTax.TextMatrix(5, 0) = "Purchase Tax:"
            grdTax.ColAlignment(1) = flexAlignRightCenter
            grdTax.TextMatrix(1, 1) = "0.00"
            grdTax.TextMatrix(2, 1) = "0.00"
            grdTax.TextMatrix(3, 1) = "0.00"
            grdTax.TextMatrix(4, 1) = "0.00"
            grdTax.TextMatrix(5, 1) = "0.00"
            
            grdStock.TextMatrix(0, 0) = "Stock"
            grdStock.TextMatrix(0, 1) = "Stock"
            grdStock.Select 0, 0, 0, 1
            grdStock.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdStock.MergeRow(0) = True
            grdStock.Cell(flexcpFontBold, 0, 0, 0, 1) = True
            grdStock.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
            grdStock.TextMatrix(1, 0) = "Opening Stock:"
            grdStock.TextMatrix(2, 0) = "-Consumption:"
            grdStock.TextMatrix(3, 0) = "-Transfer Out:"
            grdStock.TextMatrix(4, 0) = "+Transfers In:"
            grdStock.TextMatrix(5, 0) = "+Purchases:"
            grdStock.TextMatrix(6, 0) = "Closing Stock:"
            grdStock.ColAlignment(1) = flexAlignRightCenter
            grdStock.TextMatrix(1, 1) = "0.00"
            grdStock.TextMatrix(2, 1) = "0.00"
            grdStock.TextMatrix(3, 1) = "0.00"
            grdStock.TextMatrix(4, 1) = "0.00"
            grdStock.TextMatrix(5, 1) = "0.00"
            grdStock.TextMatrix(6, 1) = "0.00"
            grdStock.Select 6, 1, 6, 1
            grdStock.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdStock.Cell(flexcpBackColor, 6, 1, 6, 1) = &HE0E0E0
            grdStock.Cell(flexcpFontBold, 6, 0, 6, 1) = True
            
            grdCred.TextMatrix(0, 0) = "Debtors"
            grdCred.TextMatrix(0, 1) = "Debtors"
            grdCred.TextMatrix(0, 2) = "Creditors"
            grdCred.TextMatrix(0, 3) = "Creditors"
            grdCred.Select 0, 0, 0, 3
            grdCred.CellBorder &H808080, 0, 0, 0, 1, 0, 1
            grdCred.MergeRow(0) = True
            grdCred.Cell(flexcpFontBold, 0, 0, 0, 3) = True
            grdCred.Cell(flexcpBackColor, 0, 0, 0, 3) = &HE0E0E0
            grdCred.TextMatrix(1, 0) = "Opening Balance:"
            grdCred.TextMatrix(2, 0) = "+Sales:"
            grdCred.TextMatrix(3, 0) = "-Journals:"
            grdCred.TextMatrix(4, 0) = "-Receive on Account:"
            grdCred.TextMatrix(5, 0) = "Closing Balance:"
            grdCred.TextMatrix(1, 2) = "Opening Balance:"
            grdCred.TextMatrix(2, 2) = "+Purchases:"
            grdCred.TextMatrix(3, 2) = "-Journals:"
            grdCred.TextMatrix(4, 2) = "-Payments:"
            grdCred.TextMatrix(5, 2) = "Closing Balance:"
            grdCred.ColAlignment(1) = flexAlignRightCenter
            grdCred.TextMatrix(1, 1) = "0.00"
            grdCred.TextMatrix(2, 1) = "0.00"
            grdCred.TextMatrix(3, 1) = "0.00"
            grdCred.TextMatrix(4, 1) = "0.00"
            grdCred.TextMatrix(5, 1) = "0.00"
            grdCred.TextMatrix(1, 3) = "0.00"
            grdCred.TextMatrix(2, 3) = "0.00"
            grdCred.TextMatrix(3, 3) = "0.00"
            grdCred.TextMatrix(4, 3) = "0.00"
            grdCred.TextMatrix(5, 3) = "0.00"
            grdCred.Select 5, 1, 5, 1
            grdCred.CellBorder &H808080, 1, 1, 1, 0, 1, 1
            grdCred.Select 5, 3, 5, 3
            grdCred.CellBorder &H808080, 1, 1, 0, 0, 1, 1
            grdCred.Cell(flexcpBackColor, 5, 1, 5, 1) = &HE0E0E0
            grdCred.Cell(flexcpBackColor, 5, 3, 5, 3) = &HE0E0E0
            grdCred.Cell(flexcpBackColor, 1, 1, 4, 1) = &HE9E9E9
            grdCred.Cell(flexcpBackColor, 1, 3, 4, 3) = &HE9E9E9
            grdCred.Cell(flexcpFontBold, 5, 0, 5, 3) = True
            grdCred.ColAlignment(0) = flexAlignRightCenter
            grdCred.ColAlignment(2) = flexAlignRightCenter
            
            grdGP.Select 0, 0, 1, 0
            grdGP.CellBorder &H808080, 0, 0, 1, 0, 1, 0
            grdGP.MergeRow(0) = True
            grdGP.Cell(flexcpFontBold, 0, 0, 1, 3) = True
            grdGP.Cell(flexcpBackColor, 0, 1, 1, 1) = &HE0E0E0
            grdGP.Cell(flexcpBackColor, 0, 3, 1, 3) = &HE0E0E0
            grdGP.TextMatrix(0, 0) = "GP Percentage (%):"
            grdGP.TextMatrix(1, 0) = "GP Value:"
            grdGP.TextMatrix(0, 2) = "Customer Count:"
            grdGP.TextMatrix(1, 2) = "Spend per Head:"
            grdGP.ColAlignment(1) = flexAlignRightCenter
            grdGP.TextMatrix(0, 3) = "0"
            grdGP.TextMatrix(1, 3) = "0.00"
            grdGP.TextMatrix(0, 1) = "0%"
            grdGP.TextMatrix(1, 1) = "0.00"
            grdGP.ColAlignment(0) = flexAlignRightCenter
            grdGP.ColAlignment(2) = flexAlignRightCenter
            grdRev.SetFocus
            grdMain.Tag = ""
            Selection_Change
End Sub
Private Sub cmb2_DropButtonClick()
    If cmb3.Visible = False Then
        cmb2.ListWidth = 0
    Else
        cmb2.ListWidth = 160
    End If
End Sub
Private Sub cmb2_GotFocus()
    If picDate.Visible = True Then Selection_Change
    picDate.Visible = False
    ButtonEx1.Value = Up
End Sub



Private Sub cmb3_Change()
    Selection_Change
End Sub
Private Sub cmb3_GotFocus()
    If picDate.Visible = True Then Selection_Change
    picDate.Visible = False
    ButtonEx1.Value = Up
End Sub
Private Sub cmdMenu_Click(Index As Integer)
    On Error Resume Next
    If Buttime.Visible = True Then Buttime.Visible = False
    If Lbltime.Visible = True Then Lbltime.Visible = False
    If Lbltimeback.Visible = True Then Lbltimeback.Visible = False
    If Pictimeback.Visible = True Then Pictimeback.Visible = False
    If PicTradecomparison.Visible = True Then PicTradecomparison.Visible = False
    cmdSearch.Visible = False
    cmdPrice.Visible = fasle
    picDate.Visible = False
    picMain.Visible = False
    picAnalysis.Visible = False
    picLive.Visible = False
    DoEvents
    frmMain.Toolbar1.Buttons(2).Caption = "New"
    frmMain.Toolbar1.Buttons(2).Enabled = False
    frmMain.Toolbar1.Buttons(3).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    cmb2.Clear
    Select Case cmdMenu(Index).Caption
       Case "Time Sheets"
            cmb2.AddItem "Staff Shift Report"
            cmb2.AddItem "Staff Commision Report"
            cmb2.AddItem "Staff Wages"
            cmb2.Text = "Staff Shift Report"
            cmb2.ListWidth = 0
       Case "Purchases"
            cmb2.AddItem "Placed Purchase Orders"
            cmb2.AddItem "Cost Price Variations"
            cmb2.AddItem "Summary by Supplier"
            cmb2.AddItem "Analysis by Product"
            cmb2.Text = "Placed Purchase Orders"
            cmb2.ListWidth = 0
        Case "Rooms"
            cmb2.AddItem "Room Accounts"
            cmb2.AddItem "Room Sales"
            cmb2.AddItem "Deposits Paid"
            cmb2.AddItem "Payments Received"
            cmb2.Text = "Room Accounts"
            cmb2.ListWidth = 0
        Case "Debtors"
            cmb2.AddItem "Age Analysis"
            cmb2.AddItem "Debtor Accounts"
            cmb2.AddItem "Receive on Account"
            cmb2.Text = "Debtor Accounts"
            cmb2.ListWidth = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Debtor Reports')"
        Case "Sales"
            
            cmb2.AddItem "Daily Trade Analysis"
            cmb2.AddItem "Trade Analysis"
            cmb2.AddItem "Trade Comparison"
            cmb2.AddItem "Sales Analysis by User type"
'            cmb2.AddItem "Sales Analysis (Graph)"
            cmb2.AddItem "Sales Analysis by Product"
            cmb2.AddItem "Sales Analysis by Department"
            cmb2.AddItem "Sales Analysis by Location"
            cmb2.AddItem "Sales Analysis by Debtor"
            cmb2.AddItem "Sales Analysis by User"
            cmb2.AddItem "Sales Analysis by Supplier"
            cmb2.AddItem "Sales Analysis by Hour"
            cmb2.AddItem "Product Analysis"
            'cmb2.AddItem "Pre-Sales Analysis by Price"
            If Branch_Type = 10 Then cmb2.AddItem "Pre-Sales Analysis by Price"
            cmb2.Text = "Daily Trade Analysis"
            cmb2.ListWidth = 0
            frmMain.Toolbar1.Buttons(2).Caption = "End-of-Day"
            frmMain.Toolbar1.Buttons(2).Enabled = True
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Sales Reports')"
            DoEvents
            Selection_Change
        
        Case "Stock Takes"
            cmb2.AddItem "Stock Variance"
            cmb2.AddItem "Stock Takes"
            cmb2.Text = "Stock Takes"
            cmb2.ListWidth = 0
        Case "Stock"
            cmb2.AddItem "Stock Summary"
            cmb2.AddItem "Stock on Hand"
            cmb2.AddItem "Stock Levels Low"
            cmb2.AddItem "Stock on Hand (Suppliers)"
            cmb2.AddItem "Stock Movement (Quantities)"
            cmb2.AddItem "Stock Movement (Values)"
            cmb2.AddItem "Pack Links"
            cmb2.Text = "Stock Summary"
            cmb2.ListWidth = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Stock Reports')"
        Case "Journals"
            cmb2.AddItem "Purchase Journal"
            cmb2.AddItem "Sales Journal"
            cmb2.AddItem "Transfer Journal"
            cmb2.AddItem "Payout Journal"
            cmb2.AddItem "User Journals"
            cmb2.AddItem "Sales Corrections"
            cmb2.AddItem "Discounted Sales"
            cmb2.AddItem "Table and Tab Journal"
            cmb2.Text = "Sales Journal"
            cmb2.ListWidth = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'Journals')"
    End Select
    On Error GoTo 0
End Sub
Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
        For i = 0 To cmdMenu.Count - 1
        If Index = i Then
            cmdMenu(i).FontTextCaption.Bold = True
            cmdMenu(i).BackColor = &HFFC0C0
        Else
            cmdMenu(i).FontTextCaption.Bold = False
            cmdMenu(i).BackColor = &HFFFFFF
        End If
    Next i
End Sub

Private Sub cmdOk_Click()
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
    picDate.Visible = False
    If picDate.Visible = False Then
    'lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
    Selection_Change
    End If
End Sub

Private Sub cmdPrice_Click()
    frmReports.Tag = "1"
    Load frmPriceChange
    frmPriceChange.Tag = "CostVar"
    DoEvents
    frmPriceChange.Show vbModal
    frmMain.SetFocus
End Sub
Private Sub cmdSearch_Click()
    On Error Resume Next
    frmReports.Tag = "1"
    frmProdFind1.Show vbModal
    If Trim(Replace(ProductFilter(2), Chr(0), "")) = "" Then
        On Error GoTo 0
        Exit Sub
    End If
    grdMain.Rows = 1
    If Trim(ProductFilter(2)) <> "" Then
        If Right(Str(Time_Stop), 2) = "AM" Then
            Selender = DateAdd("d", 1, mthViewEnd.Value)
            lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(Selender, "DD MMM YYYY")
        Else
            Selender = mthViewEnd.Value
        End If
        DepString = "%"
        Suppstring = "%"
        If cmb1.Text = "<All Suppliers>" Then
            Suppstring = "%"
        Else
            If InStr(Mid(cmb1.Text, 1, InStrRev(cmb1.Text, "-") - 2), "-") = 0 Then
                Suppstring = Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1)) & "%"
            Else
                Suppstring = Trim(Mid(cmb1.Text, InStrRev(cmb1.Text, "-") + 1))
            End If
        End If
        If cmb3.Text = "<All Departments>" Then
            DeptString = "%"
        Else
            If InStr(Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2), "-") = 0 Then
                DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2) & "%"
            Else
                DeptString = Mid(cmb3.Text, 1, InStrRev(cmb3.Text, "-") - 2)
            End If
        End If
        
        ActiveReadServer "Select *,(Select Supplier_Name from Suppliers where Suppliers.Supplier_No = Cost_Change_View.Supplier_No) as Supplier from Cost_Change_View " & _
        " where Description like '%" & Trim(ProductFilter(2)) & "%' and Department_No like '" & DeptString & "' and Supplier_No like '" & Suppstring & "' and " & _
        " (Date_Time > '" & mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "') order by Description"
        grdTotal.TextMatrix(0, 0) = " Products = " & rs.RecordCount
        grdTotal.TextMatrix(0, 1) = " Products = " & rs.RecordCount
        grdTotal.TextMatrix(0, 2) = " Products = " & rs.RecordCount
        grdTotal.TextMatrix(0, 3) = " Products = " & rs.RecordCount
        grdTotal.TextMatrix(0, 4) = "0.00"
        grdTotal.TextMatrix(0, 5) = "0.00"
        grdTotal.TextMatrix(0, 6) = "0.00"
        Ave_Tot = 0
        While Not rs.EOF
            grdMain.Rows = grdMain.Rows + 1
            grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
            grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description") & ""
            grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Supplier") & " - (" & Trim(rs.Fields("Supplier_No")) & ")"
            grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("Invoice_No")
            grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("Prev_Cost"), "0.00")
            If Val(rs.Fields("Prev_Cost") & "") = 0 Then
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = "100%"
            Else
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = Round(((Val(rs.Fields("Landed_Cost") & "") - Val(rs.Fields("Prev_Cost") & "")) / Val(rs.Fields("Prev_Cost") & "")) * 100, 2) & "%"
                Ave_Tot = Ave_Tot + grdMain.ValueMatrix(grdMain.Rows - 1, 5)
            End If
            grdMain.TextMatrix(grdMain.Rows - 1, 6) = Format(rs.Fields("Landed_Cost"), "0.00")
            rs.MoveNext
        Wend
        If rs.RecordCount - 1 <> 0 Then
            grdTotal.TextMatrix(0, 4) = "Average Price Change = " & Round(Ave_Tot / rs.RecordCount - 1, 2)
            grdTotal.TextMatrix(0, 5) = "Average Price Change = " & Round(Ave_Tot / rs.RecordCount - 1, 2)
            grdTotal.TextMatrix(0, 6) = "Average Price Change = " & Round(Ave_Tot / rs.RecordCount - 1, 2)
        Else
            grdTotal.TextMatrix(0, 4) = "Average Price Change = 0"
            grdTotal.TextMatrix(0, 5) = "Average Price Change = 0"
            grdTotal.TextMatrix(0, 6) = "Average Price Change = 0"
        End If
        rs.Close
        If grdMain.Rows > 1 Then
            grdMain.Row = 1
            grdMain.SetFocus
        End If
        ProductFilter(2) = ""
    End If
    On Error GoTo 0
End Sub
Private Sub Form_Activate()
    DoEvents
    If frmReports.Tag = "1" Then
        frmReports.Tag = ""
        Exit Sub
    End If
    DoEvents
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = "Reports"
    If Buttime.Visible = True Then Buttime.Visible = False
    If Lbltime.Visible = True Then Lbltime.Visible = False
    If Lbltimeback.Visible = True Then Lbltimeback.Visible = False
    If Pictimeback.Visible = True Then Pictimeback.Visible = False
    If PicTradecomparison.Visible = True Then PicTradecomparison.Visible = False
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
Private Sub Form_Load()
    picDate.Height = 1065
    grdMain.RowHeight(0) = 400
    mthViewStart.Value = Date
    mthViewEnd.Value = Date
    frmMain.stbBar.Panels(3) = "Business Hours from " & Format(Time_Start, "hh:mm:ss AM/PM") & " to " & Format(Time_Stop, "hh:mm:ss AM/PM")
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
    
    grdTotal.MergeCells = flexMergeRestrictRows
    frmMain.Toolbar1.Buttons(4).Enabled = False
    frmMain.Toolbar1.Buttons(11).Enabled = True
    frmMain.Toolbar1.Buttons(12).Enabled = True
    cmdMenu_Click 0
    For i = 0 To cmdMenu.Count - 1
        cmdMenu(i).Value = 0
        cmdMenu(i).FontTextCaption.Bold = False
        cmdMenu(i).BackColor = &HFFFFFF
    Next i
    cmdMenu(0).Value = 1
    cmdMenu(0).FontTextCaption.Bold = True
    cmdMenu(0).BackColor = &HFFC0C0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(16).Enabled = False
    frmMain.Toolbar1.Buttons(16).Tag = ""
    frmMain.Toolbar1.Buttons(2).Caption = "New"
    
End Sub

Private Sub grdMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        Select Case cmb2.Text
            Case "Sales Journal", "Room Sales", "Sales Corrections", "Discounted Sales"
                frmTransView.LoadOldTable frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 3)
                End Select
    End If
End Sub

Private Sub grdMain_DblClick()
    Select Case cmb2.Text
        Case "Product Analysis"
            frmReports.Tag = "1"
            Load frmPriceChange
            frmPriceChange.Tag = "ProdAN"
            DoEvents
            frmPriceChange.Show vbModal
        Case "Cost Price Variations"
            frmReports.Tag = "1"
            frmPurchase.Show vbModal
        Case "Placed Purchase Orders"
            frmReports.Tag = "1"
            rptOrder.Show
        Case "Stock Takes"
            frmReports.Tag = "1"
            rptVariance.Show
        Case "Staff Commision Report"
            frmWage.Show vbModal
        Case "Room Sales"
            frmTransView.Show 0, frmMain
        Case "Sales Journal", "Sales Corrections", "Discounted Sales"
            TillData.Account_No = ""
            frmTransView.Show 0, frmMain
        Case "Purchase Journal"
            frmReports.Tag = "1"
            rptGRV.Show
        Case "Transfer Journal"
            frmReports.Tag = "1"
            If Branch_Type = 10 Then
                rptTransfersW.Show
            Else
                rptTransfers.Show
            End If
        Case "Deposits Paid", "Payments Received"
            frmReports.Tag = "1"
            TillData.Res_No = grdMain.TextMatrix(grdMain.Row, 8)
            rptRoomAcc.Show
        Case "Room Accounts"
            frmReports.Tag = "1"
            TillData.Res_No = grdMain.TextMatrix(grdMain.Row, 7)
            rptRoomAcc.Show
    End Select
End Sub
Private Sub grdMain_GotFocus()
    picDate.Visible = False
    ButtonEx1.Value = Up
End Sub
Private Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            Select Case cmb2.Text
                Case "Sales Journal"
                    frmTransView.Show 0, frmMain
                Case "Purchase Journal"
                    frmReports.Tag = "1"
                    rptGRV.Show
            End Select
    End Select
End Sub

Private Sub grdStock_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    frmGP.Show vbModal
End Sub

Private Sub grdStock_EnterCell()
    Select Case grdStock.Col
        Case 1
            Select Case grdStock.Row
                Case 2
                    grdStock.Editable = flexEDKbdMouse
                    grdStock.ComboList = "..."
                Case Else
                    grdStock.Editable = flexEDNone
                    grdStock.ComboList = ""
            End Select
        Case Else
            grdStock.Editable = flexEDNone
            grdStock.ComboList = ""
    End Select
End Sub
Private Sub mthView_LostFocus()
    DoEvents
    If picDate.Visible = False Then Selection_Change
End Sub
Private Sub mthView_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & Format(mthViewEnd.Value, "DD MMM YYYY")
End Sub



Private Sub Timer1_Timer()
Timer1.Enabled = False
Selection_Change

End Sub
