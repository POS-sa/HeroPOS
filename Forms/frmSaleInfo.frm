VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSaleInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Sales Consumption..."
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   6360
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6495
      _cx             =   11456
      _cy             =   11218
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
      GridColor       =   16777215
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSaleInfo.frx":0000
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
      ExplorerBar     =   0
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
         TabIndex        =   1
         Top             =   11700
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Left            =   5340
      TabIndex        =   2
      Top             =   6510
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx cmdOk1 
      Height          =   345
      Left            =   4050
      TabIndex        =   3
      Top             =   6510
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Open"
      Enabled         =   0   'False
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
Attribute VB_Name = "frmSaleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
    Unload Me
End Sub
Private Sub cmdOk1_Click()
    TillData.Account_No = ""
    Load frmTransView
    frmTransView.Tag = Trim(Mid(grdMain.TextMatrix(grdMain.Row, 1), 1, InStr(grdMain.TextMatrix(grdMain.Row, 1), ">") - 1))
    frmTransView.Show vbModal
End Sub
Private Sub Form_Load()
    Select Case frmInquiry.cmbLocs.Text
        Case "<All Locations>"
            LocString = "%"
        Case Else
            LocString = Trim(Mid(frmInquiry.cmbLocs.Text, 1, InStr(frmInquiry.cmbLocs.Text, "-") - 1))
    End Select
    grdMain.TextMatrix(0, 0) = "Dated"
    grdMain.TextMatrix(0, 1) = "Invoice Detail"
    grdMain.TextMatrix(0, 2) = "Qty Sold"
    grdMain.ColWidth(0) = grdMain.Width * 0.4
    grdMain.ColWidth(1) = grdMain.Width * 0.4
    grdMain.ColWidth(2) = grdMain.Width * 0.2
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignRightCenter
    grdMain.Rows = 1
    ActiveReadServer "Select sum(Qty_Consumed) as Qty,Invoice_No,Location_No,(Select Loc_Name from Locations where Consumption_Journal.Location_No = Locations.Location_No) as Location,max(Date_Time) as Date_Time from Consumption_Journal where Product_Code='" & frmInquiry.txtProdCode.Text & "' and Location_No like '" & LocString & "' and " & _
    "Date_Time > '" & frmInquiry.mthViewStart.Value & " " & Format("00:00:00", "hh:mm:ss AM/PM") & "' and Date_Time<'" & frmInquiry.mthViewEnd.Value & " " & Format("23:59:59", "hh:mm:ss AM/PM") & "' group by Invoice_No,Location_No order by Date_Time"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Format(rs.Fields("Date_Time"), "YYYY-MM-DD DDD HH:MM")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Invoice_No") & " > " & rs.Fields("Location_No") & " - " & rs.Fields("Location")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Round(rs.Fields("Qty"), 3)
        rs.MoveNext
    Wend
    rs.Close
    If grdMain.Rows > 1 Then
        cmdOk1.Enabled = True
        grdMain.Row = 1
    End If
End Sub
