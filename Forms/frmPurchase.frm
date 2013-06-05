VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPurchase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Purchase History..."
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProdCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   2265
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   570
      Width           =   7815
   End
   Begin VSFlex8Ctl.VSFlexGrid grdSuppliers 
      Height          =   2910
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   9420
      _cx             =   16616
      _cy             =   5133
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
      BackColorSel    =   16642749
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPurchase.frx":0000
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
   Begin btButtonEx.ButtonEx ButtonEx4 
      Height          =   345
      Left            =   8190
      TabIndex        =   1
      Top             =   3960
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Close"
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
   Begin btButtonEx.ButtonEx cmdOpen 
      Height          =   345
      Left            =   60
      TabIndex        =   6
      Top             =   3960
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "View GRV"
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
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9540
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code: "
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description: "
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   510
      Width           =   1335
   End
   Begin MSForms.Image Image11 
      Height          =   345
      Left            =   1470
      Top             =   90
      Width           =   2385
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "4207;609"
   End
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   1470
      Top             =   480
      Width           =   7965
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "14049;609"
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx4_Click()
    Unload Me
End Sub
Private Sub cmdOpen_Click()
    On Error Resume Next
    TillData.DocNo = grdSuppliers.TextMatrix(grdSuppliers.Row, 5)
    rptGRV1.Show vbModal
    TillData.DocNo = 0
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    Select Case frmPurchase.Tag
        Case ""
            txtProdCode.Text = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 0)
            txtDescription.Text = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 1)
        Case "GRV"
            txtProdCode.Text = frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0)
            txtDescription.Text = frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 1)
        Case "Order"
            txtProdCode.Text = frmOrder.grdOrder.TextMatrix(frmOrder.grdOrder.Row, 0)
            txtDescription.Text = frmOrder.grdOrder.TextMatrix(frmOrder.grdOrder.Row, 2)
    End Select
    LoadSuppInfo
End Sub
Private Sub Form_Load()
    grdSuppliers.Rows = 1
    grdSuppliers.TextMatrix(0, 0) = "Supplier No."
    grdSuppliers.TextMatrix(0, 1) = "Supplier Name"
    grdSuppliers.TextMatrix(0, 2) = "Qty Received"
    grdSuppliers.TextMatrix(0, 3) = "Delivered On"
    grdSuppliers.TextMatrix(0, 4) = "Landed Cost"
    grdSuppliers.ColWidth(0) = grdSuppliers.Width * 0.15
    grdSuppliers.ColWidth(1) = grdSuppliers.Width * 0.35
    grdSuppliers.ColWidth(2) = grdSuppliers.Width * 0.15
    grdSuppliers.ColAlignment(2) = flexAlignRightCenter
    grdSuppliers.ColWidth(3) = grdSuppliers.Width * 0.2
    grdSuppliers.ColWidth(4) = grdSuppliers.Width * 0.12
    grdSuppliers.ColAlignment(4) = flexAlignRightCenter
End Sub
Private Sub LoadSuppInfo()
    Screen.MousePointer = 11
    ActiveReadServer3 "SELECT Purchase_Journal.Grv_No,Purchase_Journal.Supplier_No, Suppliers.Supplier_Name, Purchase_Journal.Qty_Invoiced," & _
    " Purchase_Journal.Invoice_Date , Purchase_Journal.Price_Invoiced, Purchase_Journal.Product_Code" & _
    " FROM Purchase_Journal LEFT OUTER JOIN" & _
    " Suppliers ON Purchase_Journal.Supplier_No = Suppliers.Supplier_No" & _
    " WHERE (Purchase_Journal.Product_Code = '" & txtProdCode.Text & "')" & _
    " ORDER BY Purchase_Journal.Invoice_Date DESC"
    grdSuppliers.Rows = 1
    grdSuppliers.ColHidden(5) = True
    While Not rs3.EOF
        grdSuppliers.Rows = grdSuppliers.Rows + 1
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 0) = rs3.Fields("Supplier_No")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 1) = rs3.Fields("Supplier_Name") & ""
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 2) = rs3.Fields("Qty_Invoiced")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 3) = Format(rs3.Fields("Invoice_Date"), "YYYY-MM-DD HH:MM")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 4) = Format(rs3.Fields("Price_Invoiced"), "0.00")
        grdSuppliers.TextMatrix(grdSuppliers.Rows - 1, 5) = rs3.Fields("GRv_No")
        rs3.MoveNext
    Wend
    rs3.Close
    Screen.MousePointer = 0
End Sub
Private Sub grdSuppliers_GotFocus()
    On Error Resume Next
    If grdSuppliers.Rows > 1 Then
        cmdOpen.Enabled = True
        If grdSuppliers.Row = 0 Then grdSuppliers.Row = 1
    End If
    On Error GoTo 0
End Sub
