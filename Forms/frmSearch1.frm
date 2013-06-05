VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSearch1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Search"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "frmSearch1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   345
      Left            =   4740
      TabIndex        =   1
      Top             =   7410
      Width           =   1395
      _ExtentX        =   2461
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
   Begin VSFlex8Ctl.VSFlexGrid grdProd 
      Bindings        =   "frmSearch1.frx":000C
      Height          =   7230
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6060
      _cx             =   10689
      _cy             =   12753
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
      FormatString    =   $"frmSearch1.frx":0022
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
      Begin MSAdodcLib.Adodc adoData 
         Height          =   375
         Left            =   90
         Top             =   360
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         ConnectMode     =   1
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adoData"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin btButtonEx.ButtonEx cmdOk 
      Height          =   345
      Left            =   3270
      TabIndex        =   2
      Top             =   7410
      Width           =   1395
      _ExtentX        =   2461
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
Attribute VB_Name = "frmSearch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmSearch1.Tag = ""
    Me.Hide
End Sub
Private Sub cmdOK_Click()
    frmSearch1.Tag = grdProd.TextMatrix(grdProd.Row, 1) & " - " & grdProd.TextMatrix(grdProd.Row, 0)
    Me.Hide
End Sub
Private Sub Form_Activate()
    Static OldLocation
    DoEvents
    Screen.MousePointer = 11
    Select Case frmSearch.Tag
        Case ""
            Location = Val(Mid(frmOrder.grdOrder.TextMatrix(frmOrder.grdOrder.Row, 2), 1, InStr(frmOrder.grdOrder.TextMatrix(frmOrder.grdOrder.Row, 2), "-") - 1))
        Case Else
            Location = Location_No
    End Select
    ff = grdProd.Rows
    If Location <> OldLocation Or grdProd.Rows = 1 Then
        grdProd.Rows = 1
        DoEvents
        adoData.ConnectionString = cnnMain.ConnectionString
        adoData.CursorLocation = adUseClient
        adoData.CursorType = adOpenKeyset
        adoData.LockType = adLockBatchOptimistic
        frmSearch.Tag = ""
        adoData.ConnectionString = cnnMain.ConnectionString
        adoData.CursorLocation = adUseServer
        adoData.CursorType = adOpenStatic
        adoData.LockType = adLockBatchOptimistic
        ActiveReadServer "Select Loc_Type from Locations where Location_No = " & Location
        Select Case rs.Fields("Loc_Type")
            Case 0, 1
                adoData.RecordSource = "Select Product_Code,  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
                "+ Unit_of_Measure END AS Description from Products where Stock_Item = 1 and Department_No in" & _
                " (Select Department_No from Departments_Stock where Location_No = " & Location & ")  order by Description"
            Case 2
                adoData.RecordSource = "Select Product_Code,  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
            "+ Unit_of_Measure END AS Description from Products where Stock_Item = 1 and Department_No in" & _
            " (Select Department_No from Departments_Stock where Location_No = " & Location & ")  or (Stock_Item = 0 and Sales_Item = 0) order by Description"
        End Select
        adoData.Refresh
        grdProd.SetFocus
        grdProd.Col = 1
    End If
    Screen.MousePointer = 0
    OldLocation = Location
    grdProd.ColAlignment(0) = flexAlignLeftCenter
    grdProd.ColAlignment(1) = flexAlignLeftCenter
End Sub
Private Sub grdProd_DblClick()
    frmSearch1.Tag = grdProd.TextMatrix(grdProd.Row, 1) & " - " & grdProd.TextMatrix(grdProd.Row, 0)
    Me.Hide
End Sub

Private Sub grdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            frmSearch1.Tag = ""
            Me.Hide
        Case 13
            frmSearch1.Tag = grdProd.TextMatrix(grdProd.Row, 1) & " - " & grdProd.TextMatrix(grdProd.Row, 0)
            Me.Hide
    End Select
End Sub
