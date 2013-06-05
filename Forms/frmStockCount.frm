VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmStockCount 
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmStockCount.frx":0000
   ScaleHeight     =   7530
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   840
      Index           =   11
      Left            =   4050
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1482
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
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   5985
      Left            =   300
      TabIndex        =   2
      Top             =   1260
      Width           =   4635
      _cx             =   8176
      _cy             =   10557
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   15523287
      ForeColorSel    =   0
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
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   3300
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
   Begin VB.Label lblHeading 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   510
      TabIndex        =   1
      Top             =   510
      Width           =   2715
   End
End
Attribute VB_Name = "frmStockCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKey_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    grdMain.TextMatrix(0, 0) = "Description"
    grdMain.TextMatrix(0, 1) = "Count"
    lblHeading = ""
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignRightCenter
    grdMain.Rows = 1
    For i = 1 To frmStockTake.grdMain.Rows - 1
        If frmStockTake.grdMain.TextMatrix(i, 2) = frmStockTake.grdMerge.TextMatrix(frmStockTake.grdMerge.Row, 4) Then
            grdMain.Rows = grdMain.Rows + 1
            grdMain.TextMatrix(grdMain.Rows - 1, 0) = frmStockTake.grdMain.TextMatrix(i, 0)
            grdMain.TextMatrix(grdMain.Rows - 1, 1) = frmStockTake.grdMain.TextMatrix(i, 1)
        End If
    Next i
    Select Case grdMain.Rows
        Case 2
            lblHeading = "Counted Once"
        Case Else
            lblHeading = "Counted " & grdMain.Rows - 1 & " Times"
    End Select
End Sub
