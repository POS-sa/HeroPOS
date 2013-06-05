VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDebTrans 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Transaction Transfer..."
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   345
      Left            =   3510
      TabIndex        =   0
      Top             =   5100
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx cmdOK 
      Height          =   345
      Left            =   2340
      TabIndex        =   1
      Top             =   5100
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Ok"
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
   Begin VSFlex8Ctl.VSFlexGrid grdGrid 
      Height          =   4950
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4545
      _cx             =   8017
      _cy             =   8731
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
      BackColorSel    =   15329975
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDebTrans.frx":0000
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
      Editable        =   2
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
Attribute VB_Name = "frmDebTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmDebTrans.Tag = ""
    Me.Hide
End Sub
Private Sub cmdOK_Click()
    frmDebTrans.Tag = grdGrid.TextMatrix(grdGrid.Row, 0)
    Me.Hide
End Sub

Private Sub Form_Load()
    frmDebTrans.Tag = ""
    grdGrid.Rows = 1
    grdGrid.TextMatrix(0, 0) = "Debtor No"
    grdGrid.TextMatrix(0, 1) = "Debtor Name"
    grdGrid.ColAlignment(0) = flexAlignLeftCenter
    grdGrid.ColAlignment(1) = flexAlignLeftCenter
    grdGrid.ColWidth(0) = grdGrid.Width * 0.3
    grdGrid.ColWidth(1) = grdGrid.Width * 0.7
    ActiveReadServer "Select Debtor_No,Debtor_Name from Debtors order by Debtor_No"
    While Not rs.EOF
        grdGrid.Rows = grdGrid.Rows + 1
        grdGrid.TextMatrix(grdGrid.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = rs.Fields("Debtor_Name")
        rs.MoveNext
    Wend
    rs.Close
End Sub
