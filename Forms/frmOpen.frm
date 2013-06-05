VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Document"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   7230
      TabIndex        =   0
      ToolTipText     =   " Click to Search.... "
      Top             =   4410
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdOpen 
      Height          =   375
      Left            =   5940
      TabIndex        =   1
      ToolTipText     =   " Click to Search.... "
      Top             =   4410
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      ShowFocus       =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4245
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   7488
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Saved GRV's"
      TabPicture(0)   =   "frmOpen.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdList(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Saved Orders"
      TabPicture(1)   =   "frmOpen.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdList(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Orders"
      TabPicture(2)   =   "frmOpen.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "grdList(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VSFlex8Ctl.VSFlexGrid grdList 
         Height          =   3750
         Index           =   0
         Left            =   -74910
         TabIndex        =   3
         Top             =   390
         Width           =   8175
         _cx             =   14420
         _cy             =   6615
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         BackColorSel    =   15391677
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOpen.frx":0060
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
      End
      Begin VSFlex8Ctl.VSFlexGrid grdList 
         Height          =   3750
         Index           =   1
         Left            =   -74910
         TabIndex        =   4
         Top             =   390
         Width           =   8175
         _cx             =   14420
         _cy             =   6615
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         BackColorSel    =   15391677
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOpen.frx":00D8
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
      End
      Begin VSFlex8Ctl.VSFlexGrid grdList 
         Height          =   3750
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   390
         Width           =   8175
         _cx             =   14420
         _cy             =   6615
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         BackColorSel    =   15391677
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16645618
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOpen.frx":0150
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
      End
   End
   Begin btButtonEx.ButtonEx cmdDelete 
      Height          =   375
      Left            =   60
      TabIndex        =   6
      ToolTipText     =   " Click to Search.... "
      Top             =   4410
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Delete"
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
      ShowFocus       =   0
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    Select Case SSTab1.Tab
        Case 0
            If MsgBox("Are you sure you want to delete this Saved GRV?", vbYesNo, "HeroPOS") = vbYes Then
                ActiveUpdateServer "Delete from Purchase_Journal_Listing where Grv_No = " & Val(grdList(0).TextMatrix(grdList(0).Row, 2))
                DoEvents
                ActiveReadServer "Select * from GRV_Listing order by Grv_No"
                grdList(0).Rows = 1
                While Not rs.EOF
                    grdList(0).Rows = grdList(0).Rows + 1
                    grdList(0).TextMatrix(grdList(0).Rows - 1, 0) = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
                    grdList(0).TextMatrix(grdList(0).Rows - 1, 1) = rs.Fields("User_no") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
                    grdList(0).TextMatrix(grdList(0).Rows - 1, 2) = Format(rs.Fields("Grv_no"), "000000")
                    grdList(0).TextMatrix(grdList(0).Rows - 1, 3) = rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
                    rs.MoveNext
                Wend
                rs.Close
            End If
            grdList(0).Row = 0
        Case 1
            If MsgBox("Are you sure you want to delete this Saved Order?", vbYesNo, "HeroPOS") = vbYes Then
                ActiveUpdateServer "Delete from Purchase_Order_Listing where Order_No = " & Val(grdList(1).TextMatrix(grdList(1).Row, 2))
                DoEvents
                ActiveReadServer "Select * from Order_Listing order by Order_No"
                grdList(1).Rows = 1
                While Not rs.EOF
                    grdList(1).Rows = grdList(1).Rows + 1
                    grdList(1).TextMatrix(grdList(1).Rows - 1, 0) = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
                    grdList(1).TextMatrix(grdList(1).Rows - 1, 1) = rs.Fields("User_no") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
                    grdList(1).TextMatrix(grdList(1).Rows - 1, 2) = Format(rs.Fields("Order_no"), "000000")
                    grdList(1).TextMatrix(grdList(1).Rows - 1, 3) = rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
                    rs.MoveNext
                Wend
                rs.Close
            End If
            grdList(1).Row = 0
        Case 2
            If MsgBox("Are you sure you want to delete this Placed Order?", vbYesNo, "HeroPOS") = vbYes Then
                ActiveUpdateServer "Delete from Purchase_Order_Journal where Order_No = " & Val(grdList(2).TextMatrix(grdList(2).Row, 2))
                DoEvents
                ActiveReadServer "Select * from Placed_Order_Listing order by Order_No"
                grdList(2).Rows = 1
                While Not rs.EOF
                    grdList(2).Rows = grdList(2).Rows + 1
                    grdList(2).TextMatrix(grdList(2).Rows - 1, 0) = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
                    grdList(2).TextMatrix(grdList(2).Rows - 1, 1) = rs.Fields("User_no") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
                    grdList(2).TextMatrix(grdList(2).Rows - 1, 2) = Format(rs.Fields("Order_no"), "000000")
                    grdList(2).TextMatrix(grdList(2).Rows - 1, 3) = rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
                    rs.MoveNext
                Wend
                rs.Close
            End If
            grdList(2).Row = 0
    End Select
End Sub
Private Sub cmdOpen_Click()
    Select Case SSTab1.Tab
        Case 0
            frmGRV.LoadGrv Val(grdList(0).TextMatrix(grdList(0).Row, 2))
        Case 1
            frmOrder.LoadOrder Val(grdList(1).TextMatrix(grdList(1).Row, 2))
        Case 2
            frmGRV.LoadOrder Val(grdList(2).TextMatrix(grdList(2).Row, 2))
        Case Else
            frmGRV.LoadGrv Val(grdList(0).TextMatrix(grdList(0).Row, 2))
    End Select
    Unload Me
End Sub
Private Sub Form_Activate()
    ActiveReadServer "Select * from GRV_Listing order by Grv_No"
    grdList(0).Rows = 1
    While Not rs.EOF
        grdList(0).Rows = grdList(0).Rows + 1
        grdList(0).TextMatrix(grdList(0).Rows - 1, 0) = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
        grdList(0).TextMatrix(grdList(0).Rows - 1, 1) = rs.Fields("User_no") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
        grdList(0).TextMatrix(grdList(0).Rows - 1, 2) = Format(rs.Fields("Grv_no"), "000000")
        grdList(0).TextMatrix(grdList(0).Rows - 1, 3) = rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
        rs.MoveNext
    Wend
    rs.Close
    If grdList(0).TextMatrix(grdList(0).Row, 0) = "" Or Trim(grdList(0).TextMatrix(grdList(0).Row, 0)) = "Dated" Then
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdOpen.Enabled = True
        cmdDelete.Enabled = True
    End If
    ActiveReadServer "Select * from Order_Listing order by Order_No"
    grdList(1).Rows = 1
    While Not rs.EOF
        grdList(1).Rows = grdList(1).Rows + 1
        grdList(1).TextMatrix(grdList(1).Rows - 1, 0) = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
        grdList(1).TextMatrix(grdList(1).Rows - 1, 1) = rs.Fields("User_no") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
        grdList(1).TextMatrix(grdList(1).Rows - 1, 2) = Format(rs.Fields("Order_no"), "000000")
        grdList(1).TextMatrix(grdList(1).Rows - 1, 3) = rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
        rs.MoveNext
    Wend
    rs.Close
    If grdList(1).TextMatrix(grdList(1).Row, 0) = "" Or Trim(grdList(1).TextMatrix(grdList(1).Row, 0)) = "Dated" Then
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdOpen.Enabled = True
        cmdDelete.Enabled = True
    End If
    ActiveReadServer "Select * from Placed_Order_Listing order by Order_No"
    grdList(2).Rows = 1
    While Not rs.EOF
        grdList(2).Rows = grdList(2).Rows + 1
        grdList(2).TextMatrix(grdList(2).Rows - 1, 0) = Format(rs.Fields("Date_Time"), "ddd dd MMM yyyy")
        grdList(2).TextMatrix(grdList(2).Rows - 1, 1) = rs.Fields("User_no") & " - " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
        grdList(2).TextMatrix(grdList(2).Rows - 1, 2) = Format(rs.Fields("Order_no"), "000000")
        grdList(2).TextMatrix(grdList(2).Rows - 1, 3) = rs.Fields("Supplier_Name") & " - " & rs.Fields("Supplier_No")
        rs.MoveNext
    Wend
    rs.Close
    If grdList(2).TextMatrix(grdList(2).Row, 0) = "" Or Trim(grdList(2).TextMatrix(grdList(2).Row, 0)) = "Dated" Then
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdOpen.Enabled = True
        cmdDelete.Enabled = True
    End If
    Select Case frmOpen.Tag
        Case "GRV"
            SSTab1.TabVisible(0) = True
            SSTab1.TabVisible(1) = False
            SSTab1.TabVisible(2) = True
        Case "Order"
            SSTab1.TabVisible(0) = False
            SSTab1.TabVisible(1) = True
            SSTab1.TabVisible(2) = False
    End Select
End Sub
Private Sub Form_Load()
    For i = 0 To 2
        grdList(i).Cols = 4
        grdList(i).ColWidth(0) = grdList(i).Width * 0.25
        grdList(i).ColWidth(1) = grdList(i).Width * 0.3
        grdList(i).ColWidth(2) = grdList(i).Width * 0.12
        grdList(i).ColWidth(3) = grdList(i).Width * 0.33
        grdList(i).TextMatrix(0, 0) = "Dated"
        grdList(i).TextMatrix(0, 1) = "User"
        grdList(i).TextMatrix(0, 2) = "Doc No"
        grdList(i).TextMatrix(0, 3) = "Supplier"
        grdList(i).ColAlignment(2) = flexAlignRightBottom
        grdList(i).ColAlignment(0) = flexAlignLeftCenter
        grdList(i).ColAlignment(1) = flexAlignLeftCenter
        grdList(i).ColAlignment(2) = flexAlignLeftCenter
        grdList(i).ColAlignment(3) = flexAlignLeftCenter
    Next i
End Sub

Private Sub grdList_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If grdList(Index).TextMatrix(grdList(Index).Row, 0) = "" Or Trim(grdList(Index).TextMatrix(grdList(Index).Row, 0)) = "Dated" Then
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdOpen.Enabled = True
        cmdDelete.Enabled = True
    End If
End Sub

Private Sub grdList_DblClick(Index As Integer)
    Select Case Index
        Case 0
            frmGRV.LoadGrv Val(grdList(0).TextMatrix(grdList(0).Row, 2))
        Case 1
            frmOrder.LoadOrder Val(grdList(1).TextMatrix(grdList(1).Row, 2))
        Case 2
            frmGRV.LoadOrder Val(grdList(2).TextMatrix(grdList(2).Row, 2))
        Case Else
            frmGRV.LoadGrv Val(grdList(0).TextMatrix(grdList(0).Row, 2))
    End Select
    Unload Me
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    cmdOpen.Enabled = False
    cmdDelete.Enabled = False
    Select Case SSTab1.Caption
        Case "Saved GRV's"
            grdList(0).Row = 0
        Case "Saved Orders"
            grdList(1).Row = 0
        Case "Orders"
            grdList(2).Row = 0
    End Select
End Sub
