VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSearchRes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Guest Search..."
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid grdRes 
      Height          =   3420
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8865
      _cx             =   15637
      _cy             =   6032
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmShearchRes.frx":0000
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   2
         Top             =   11700
         Width           =   1005
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   1
         Top             =   5670
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Index           =   0
      Left            =   7680
      TabIndex        =   3
      Top             =   3540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Cancel"
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
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Index           =   1
      Left            =   6390
      TabIndex        =   4
      Top             =   3540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
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
End
Attribute VB_Name = "frmSearchRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            frmRes.FindMe grdRes.ValueMatrix(grdRes.Row, 0)
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
    grdRes.Rows = 1
    ActiveReadServer "Select Tel_No, Last_Name,First_Name,Res_No from Reservations where Res_Type <> 3 order by Res_No"
    While Not rs.EOF
        grdRes.Rows = grdRes.Rows + 1
        grdRes.TextMatrix(grdRes.Rows - 1, 0) = rs.Fields("Res_No")
        grdRes.TextMatrix(grdRes.Rows - 1, 1) = rs.Fields("First_Name")
        grdRes.TextMatrix(grdRes.Rows - 1, 2) = rs.Fields("Last_Name")
        grdRes.TextMatrix(grdRes.Rows - 1, 3) = rs.Fields("Tel_No")
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub Form_Load()
    grdRes.TextMatrix(0, 0) = "Res No"
    grdRes.TextMatrix(0, 1) = "First Name"
    grdRes.TextMatrix(0, 2) = "Last Name"
    grdRes.TextMatrix(0, 3) = "Telephone"
    grdRes.ColWidth(0) = grdRes.Width * 0.15
    grdRes.ColWidth(1) = grdRes.Width * 0.35
    grdRes.ColWidth(2) = grdRes.Width * 0.35
    grdRes.ColWidth(3) = grdRes.Width * 0.15
    grdRes.ColAlignment(0) = flexAlignLeftCenter
    grdRes.ColAlignment(1) = flexAlignLeftCenter
    grdRes.ColAlignment(2) = flexAlignLeftCenter
    grdRes.ColAlignment(3) = flexAlignLeftCenter
    frmPrev.grdRes.Rows = 1
End Sub
Private Sub grdRes_DblClick()
    frmRes.FindMe grdRes.ValueMatrix(grdRes.Row, 0)
    Unload Me
End Sub
Private Sub grdRes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        frmRes.FindMe grdRes.ValueMatrix(grdRes.Row, 0)
        Unload Me
    End If
End Sub
