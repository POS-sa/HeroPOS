VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSim 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Find Similar Products..."
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8Ctl.VSFlexGrid grdProd 
      Bindings        =   "frmSim.frx":0000
      Height          =   4440
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3270
      _cx             =   5768
      _cy             =   7832
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSim.frx":0016
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
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   345
      Index           =   0
      Left            =   2100
      TabIndex        =   1
      Top             =   4560
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
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   345
      Index           =   1
      Left            =   810
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
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
Attribute VB_Name = "frmSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub
Private Sub cmdClose_Click(Index As Integer)
    qText = "Select Product_Code,Selling_Price, CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
                "+ Unit_of_Measure END as Description from Products where " & "CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) + Unit_of_Measure END like "
    firstSelect = False
    For i = 1 To grdProd.Rows - 1
        If grdProd.ValueMatrix(i, 1) = -1 Then
            If firstSelect = False Then
                qText = qText & "'%" & grdProd.TextMatrix(i, 0) & "%'"
                firstSelect = True
            Else
                qText = qText & " and CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size)+ Unit_of_Measure END" & " like '%" & grdProd.TextMatrix(i, 0) & "%'"
            End If
        End If
    Next i
    
    ActiveReadServer qText & " Order by Description"
    If rs.RecordCount > 0 Then
        frmPriceChange.grdGrid.Rows = 1
    End If
    While Not rs.EOF
        frmPriceChange.grdGrid.Rows = frmPriceChange.grdGrid.Rows + 1
        frmPriceChange.grdGrid.TextMatrix(frmPriceChange.grdGrid.Rows - 1, 0) = rs.Fields("Product_Code")
        frmPriceChange.grdGrid.TextMatrix(frmPriceChange.grdGrid.Rows - 1, 1) = rs.Fields("Description")
        frmPriceChange.grdGrid.TextMatrix(frmPriceChange.grdGrid.Rows - 1, 2) = Format(rs.Fields("Selling_Price"), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    Unload Me
End Sub
Private Sub Form_Activate()
    Dim AllColumns As Variant
    grdProd.Rows = 1
    AllColumns = Split(frmPriceChange.txtDescription, " ")
    For i = 0 To UBound(AllColumns)
        grdProd.Rows = grdProd.Rows + 1
        grdProd.TextMatrix(grdProd.Rows - 1, 0) = Trim(AllColumns(i))
    Next i
    grdProd.SetFocus
End Sub
Private Sub Form_Load()
    grdProd.ColAlignment(0) = flexAlignLeftCenter
    grdProd.ColAlignment(1) = flexAlignCenterCenter
    grdProd.ColWidth(0) = grdProd.Width * 0.7
    grdProd.ColWidth(1) = grdProd.Width * 0.25
    grdProd.TextMatrix(0, 0) = "Keyword"
    grdProd.TextMatrix(0, 1) = "Use"
    grdProd.ColDataType(1) = flexDTBoolean
    grdProd.Editable = flexEDKbdMouse
End Sub
