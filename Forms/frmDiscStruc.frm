VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDiscStruc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Discount Structure..."
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmDiscStruc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDepartments 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   0
      ScaleHeight     =   3825
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   8415
      Begin VSFlex8Ctl.VSFlexGrid grdMinor1 
         Height          =   3810
         Left            =   5340
         TabIndex        =   1
         Top             =   0
         Width           =   3045
         _cx             =   5371
         _cy             =   6720
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   0   'False
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
         HighLight       =   2
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
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   2000
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDiscStruc.frx":000C
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
      Begin VSFlex8Ctl.VSFlexGrid grdSub1 
         Height          =   3810
         Left            =   2580
         TabIndex        =   2
         Top             =   0
         Width           =   2805
         _cx             =   4948
         _cy             =   6720
         Appearance      =   0
         BorderStyle     =   1
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   700
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDiscStruc.frx":0084
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
      End
      Begin VSFlex8Ctl.VSFlexGrid grdMajor1 
         Height          =   3810
         Left            =   60
         TabIndex        =   3
         Top             =   0
         Width           =   3105
         _cx             =   5477
         _cy             =   6720
         Appearance      =   0
         BorderStyle     =   1
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
         ColWidthMin     =   700
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDiscStruc.frx":00FC
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
      End
   End
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   345
      Left            =   7200
      TabIndex        =   4
      Top             =   3990
      Width           =   1185
      _ExtentX        =   2090
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
   Begin btButtonEx.ButtonEx cmdOk 
      Height          =   345
      Left            =   5970
      TabIndex        =   5
      Top             =   3990
      Width           =   1185
      _ExtentX        =   2090
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
   Begin VSFlex8Ctl.VSFlexGrid grdGrid 
      Height          =   3810
      Left            =   60
      TabIndex        =   6
      Top             =   4410
      Width           =   8310
      _cx             =   14658
      _cy             =   6720
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
      BackColorSel    =   16706532
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDiscStruc.frx":0174
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
   Begin btButtonEx.ButtonEx cmdApply 
      Height          =   345
      Left            =   5460
      TabIndex        =   7
      Top             =   8340
      Width           =   2905
      _ExtentX        =   5133
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Apply Structure to Selected Debtors"
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
   Begin btButtonEx.ButtonEx cmdNone 
      Height          =   345
      Left            =   1290
      TabIndex        =   8
      Top             =   8310
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Clear All"
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
   Begin btButtonEx.ButtonEx cmdAll 
      Height          =   345
      Left            =   60
      TabIndex        =   9
      Top             =   8310
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Select All"
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
Attribute VB_Name = "frmDiscStruc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAll_Click()
    For i = 1 To grdGrid.Rows - 1
        grdGrid.TextMatrix(i, 2) = 1
    Next i
End Sub

Private Sub cmdApply_Click()
    Screen.MousePointer = 11
    For i = 1 To grdGrid.Rows - 1
        If grdGrid.ValueMatrix(i, 2) = True Then
            ActiveUpdateServer "Delete from Debtor_Discounts where Debtor_No = '" & grdGrid.TextMatrix(i, 0) & "'"
            ActiveUpdateServer "Insert Into Debtor_Discounts SELECT Department_No, '" & grdGrid.TextMatrix(i, 0) & "' AS Debtor_No, Cost_Disc, Sell_Disc, Sell_Cost, Selling_Price From Debtor_Discounts WHERE (Debtor_No = '" & frmGuests.txtSuppNo & "')"
        End If
    Next i
    For i = 1 To grdGrid.Rows - 1
        grdGrid.TextMatrix(i, 2) = 0
    Next i
    Screen.MousePointer = 0
    MsgBox "Structure Change Complete", vbInformation, "HeroPOS"
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNone_Click()
    For i = 1 To grdGrid.Rows - 1
        grdGrid.TextMatrix(i, 2) = 0
    Next i
End Sub
Private Sub cmdOk_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    grdGrid.Cols = 3
    grdGrid.Rows = 1
    grdGrid.TextMatrix(0, 0) = "Debtor No"
    grdGrid.TextMatrix(0, 1) = "Debtor Name"
    grdGrid.TextMatrix(0, 2) = "Select"
    grdGrid.ColWidth(0) = grdGrid.Width * 0.2
    grdGrid.ColWidth(1) = grdGrid.Width * 0.65
    grdGrid.ColAlignment(0) = flexAlignLeftCenter
    grdGrid.ColAlignment(1) = flexAlignLeftCenter
    grdGrid.ColAlignment(2) = flexAlignCenterCenter
    grdGrid.ColDataType(2) = flexDTBoolean
    grdMajor1.SetFocus
    If grdMajor1.Rows > 1 Then
        grdMajor1.Row = 1
        grdSub1.Rows = 1
        ActiveReadServer1 "Select * From Departments where Dept_Type=1 and Dept_Parent='" & grdMajor1.TextMatrix(grdMajor1.Row, 0) & "'"
        i = 0
        While Not rs1.EOF
            i = i + 1
            grdSub1.Rows = grdSub1.Rows + 1
            grdSub1.TextMatrix(i, 0) = rs1.Fields("Department_No")
            grdSub1.TextMatrix(i, 1) = rs1.Fields("Dept_Name")
            ActiveReadServer2 "Select * from Debtor_Discounts where Debtor_No= '" & frmGuests.txtSuppNo & "' and Department_No = '" & rs1.Fields("Department_No") & "'"
            If rs2.RecordCount > 0 Then
                grdSub1.TextMatrix(i, 2) = Format(Val(rs2.Fields("Cost_Disc") & ""), "0.00")
                grdSub1.TextMatrix(i, 3) = Format(Val(rs2.Fields("Sell_Disc") & ""), "0.00")
                grdSub1.TextMatrix(i, 4) = rs2.Fields("Sell_Cost") & ""
                If grdSub1.TextMatrix(i, 4) = "" Then grdSub1.TextMatrix(i, 4) = "No"
                Select Case Val(rs2.Fields("Selling_Price") & "")
                    Case 0: grdSub1.TextMatrix(i, 5) = "<None>"
                    Case 1: grdSub1.TextMatrix(i, 5) = "Price1"
                    Case 2: grdSub1.TextMatrix(i, 5) = "Price2"
                    Case 3: grdSub1.TextMatrix(i, 5) = "Price3"
                    Case 4: grdSub1.TextMatrix(i, 5) = "Price4"
                    Case 5: grdSub1.TextMatrix(i, 5) = "Price5"
                    Case 6: grdSub1.TextMatrix(i, 5) = "Price6"
                End Select
            Else
                grdSub1.TextMatrix(i, 2) = "0.00"
                grdSub1.TextMatrix(i, 3) = "0.00"
                grdSub1.TextMatrix(i, 4) = "No"
                grdSub1.TextMatrix(i, 5) = "<None>"
            End If
            rs2.Close
            rs1.MoveNext
        Wend
        rs1.Close
    End If
    ActiveReadServer "Select Debtor_No,Debtor_Name from Debtors where Debtor_No <> '" & frmGuests.txtSuppNo & "' order by Debtor_No"
    While Not rs.EOF
        grdGrid.Rows = grdGrid.Rows + 1
        grdGrid.TextMatrix(grdGrid.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = rs.Fields("Debtor_Name")
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub Form_Load()
    grdMinor1.Cols = 2
    grdSub1.ColHidden(2) = True
    grdSub1.ColHidden(3) = True
    grdSub1.ColHidden(4) = True
    grdSub1.ColHidden(5) = True
    picDepartments.Visible = True
    grdMajor1.TextMatrix(0, 0) = "No."
    grdMajor1.TextMatrix(0, 1) = "Major Department"
    grdSub1.TextMatrix(0, 0) = "No."
    grdSub1.TextMatrix(0, 1) = "Sub Department"
    grdMinor1.TextMatrix(0, 0) = "Type"
    grdMinor1.TextMatrix(0, 1) = "Value"
    grdMajor1.ColAlignment(0) = flexAlignLeftCenter
    grdMajor1.ColAlignment(1) = flexAlignLeftCenter
    grdSub1.ColAlignment(0) = flexAlignLeftCenter
    grdSub1.ColAlignment(1) = flexAlignLeftCenter
    grdMinor1.ColAlignment(0) = flexAlignLeftCenter
    grdMinor1.ColAlignment(1) = flexAlignRightCenter
    grdMajor1.Rows = 1
    grdSub1.Rows = 1
    grdMinor1.Rows = 5
    
    grdMinor1.TextMatrix(1, 0) = "Markup from Cost"
    grdMinor1.TextMatrix(2, 0) = "Discount on Selling"
    grdMinor1.TextMatrix(3, 0) = "Sell at Cost"
    grdMinor1.TextMatrix(4, 0) = "Use Selling Price"
    
    grdMinor1.TextMatrix(1, 1) = "0.00"
    grdMinor1.TextMatrix(2, 1) = "0.00"
    grdMinor1.TextMatrix(3, 1) = "No"
    grdMinor1.TextMatrix(4, 1) = "<None>"
    
    If grdMinor1.Rows > 0 Then grdMinor1.Row = 1
    ActiveReadServer "Select * From Departments where Dept_Type=0 order by Department_No"
    i = 0
    While Not rs.EOF
        grdMajor1.Rows = grdMajor1.Rows + 1
        i = i + 1
        grdMajor1.TextMatrix(i, 0) = rs.Fields("Department_No")
        grdMajor1.TextMatrix(i, 1) = rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    If grdMajor1.Rows > 1 Then
        grdSub1.Rows = 1
        ActiveReadServer "Select * From Departments where Dept_Type=1 and Dept_Parent= '" & grdMajor1.TextMatrix(1, 0) & "' order by Department_No"
        i = 0
        While Not rs.EOF
            grdSub1.Rows = grdSub1.Rows + 1
            i = i + 1
            grdSub1.TextMatrix(i, 0) = rs.Fields("Department_No")
            grdSub1.TextMatrix(i, 1) = rs.Fields("Dept_Name")
            rs.MoveNext
        Wend
        rs.Close
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ActiveReadServer1 "Select * from Debtor_Discounts where Debtor_No = '" & frmGuests.txtSuppNo & "'"
    If rs1.RecordCount > 0 Then
        frmGuests.txtVat.Text = "Set Per Department"
    Else
        frmGuests.txtVat.Text = "<Not Set>"
    End If
    rs1.Close
End Sub

Private Sub grdGrid_EnterCell()
    grdGrid.Editable = flexEDNone
    If grdGrid.Col = 2 Then grdGrid.Editable = flexEDKbdMouse
End Sub

Private Sub grdMajor1_Click()
    grdSub1.Rows = 1
    ActiveReadServer1 "Select * From Departments where Dept_Type=1 and Dept_Parent='" & grdMajor1.TextMatrix(grdMajor1.Row, 0) & "'"
    i = 0
    While Not rs1.EOF
        i = i + 1
        grdSub1.Rows = grdSub1.Rows + 1
        grdSub1.TextMatrix(i, 0) = rs1.Fields("Department_No")
        grdSub1.TextMatrix(i, 1) = rs1.Fields("Dept_Name")
        ActiveReadServer2 "Select * from Debtor_Discounts where Debtor_No= '" & frmGuests.txtSuppNo & "' and Department_No = '" & rs1.Fields("Department_No") & "'"
        If rs2.RecordCount > 0 Then
            grdSub1.TextMatrix(i, 2) = Format(Val(rs2.Fields("Cost_Disc") & ""), "0.00")
            grdSub1.TextMatrix(i, 3) = Format(Val(rs2.Fields("Sell_Disc") & ""), "0.00")
            grdSub1.TextMatrix(i, 4) = rs2.Fields("Sell_Cost") & ""
            If grdSub1.TextMatrix(i, 4) = "" Then grdSub1.TextMatrix(i, 4) = "No"
            grdSub1.TextMatrix(i, 5) = "Price" & rs2.Fields("Selling_Price")
            If grdSub1.TextMatrix(i, 5) = "Price" Or grdSub1.TextMatrix(i, 5) = "0" Then grdSub1.TextMatrix(i, 5) = "<None>"
        Else
            grdSub1.TextMatrix(i, 2) = "0.00"
            grdSub1.TextMatrix(i, 3) = "0.00"
            grdSub1.TextMatrix(i, 4) = "No"
            grdSub1.TextMatrix(i, 5) = "<None>"
        End If
        rs2.Close
        rs1.MoveNext
    Wend
    grdMinor1.TextMatrix(2, 1) = "0.00"
    grdMinor1.TextMatrix(1, 1) = "0.00"
    grdMinor1.TextMatrix(3, 1) = "No"
    grdMinor1.TextMatrix(4, 1) = "<None>"
    rs1.Close
End Sub

Private Sub grdMinor1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If grdMinor1.ValueMatrix(Row, Col) > 100 And Row <> 3 Then
        grdMinor1.TextMatrix(Row, Col) = 100
    End If
    grdMinor1.TextMatrix(Row, Col) = Format(grdMinor1.TextMatrix(Row, Col), "0.00")
    Select Case Row
        Case 1: grdMinor1.TextMatrix(2, Col) = "0.00": grdMinor1.TextMatrix(3, Col) = "No": grdMinor1.TextMatrix(4, Col) = "<None>"
        Case 2: grdMinor1.TextMatrix(1, Col) = "0.00": grdMinor1.TextMatrix(3, Col) = "No": grdMinor1.TextMatrix(4, Col) = "<None>"
        Case 3: grdMinor1.TextMatrix(1, Col) = "0.00": grdMinor1.TextMatrix(2, Col) = "0.00": grdMinor1.TextMatrix(4, Col) = "<None>"
        Case 4: grdMinor1.TextMatrix(1, Col) = "0.00": grdMinor1.TextMatrix(2, Col) = "0.00": grdMinor1.TextMatrix(3, Col) = "No"
    End Select
    ActiveUpdateServer "Delete from Debtor_Discounts where Debtor_No= '" & frmGuests.txtSuppNo & "' and Department_No = '" & grdSub1.TextMatrix(grdSub1.Row, 0) & "'"
    Select Case Row
        Case 1
            ActiveUpdateServer "Insert Into Debtor_Discounts (Department_No,Debtor_No,Cost_Disc,Sell_Disc,Sell_Cost,Selling_Price) values ('" & grdSub1.TextMatrix(grdSub1.Row, 0) & "','" & frmGuests.txtSuppNo & "'," & Val(grdMinor1.TextMatrix(1, Col)) & "," & Val(grdMinor1.TextMatrix(2, Col)) & ",'" & grdMinor1.TextMatrix(3, Col) & "'," & Val(Right(grdMinor1.TextMatrix(4, Col), 1)) & ")"
        Case 2
            ActiveUpdateServer "Insert Into Debtor_Discounts (Department_No,Debtor_No,Cost_Disc,Sell_Disc,Sell_Cost,Selling_Price) values ('" & grdSub1.TextMatrix(grdSub1.Row, 0) & "','" & frmGuests.txtSuppNo & "'," & Val(grdMinor1.TextMatrix(1, Col)) & "," & Val(grdMinor1.TextMatrix(2, Col)) & ",'" & grdMinor1.TextMatrix(3, Col) & "'," & Val(Right(grdMinor1.TextMatrix(4, Col), 1)) & ")"
        Case 3
            ActiveUpdateServer "Insert Into Debtor_Discounts (Department_No,Debtor_No,Cost_Disc,Sell_Disc,Sell_Cost,Selling_Price) values ('" & grdSub1.TextMatrix(grdSub1.Row, 0) & "','" & frmGuests.txtSuppNo & "'," & Val(grdMinor1.TextMatrix(1, Col)) & "," & Val(grdMinor1.TextMatrix(2, Col)) & ",'" & grdMinor1.TextMatrix(3, Col) & "'," & Val(Right(grdMinor1.TextMatrix(4, Col), 1)) & ")"
        Case 4
            ActiveUpdateServer "Insert Into Debtor_Discounts (Department_No,Debtor_No,Cost_Disc,Sell_Disc,Sell_Cost,Selling_Price) values ('" & grdSub1.TextMatrix(grdSub1.Row, 0) & "','" & frmGuests.txtSuppNo & "'," & Val(grdMinor1.TextMatrix(1, Col)) & "," & Val(grdMinor1.TextMatrix(2, Col)) & ",'" & grdMinor1.TextMatrix(3, Col) & "'," & Val(Right(grdMinor1.TextMatrix(4, Col), 1)) & ")"
    End Select
    grdSub1.TextMatrix(grdSub1.Row, 2) = Format(Val(grdMinor1.TextMatrix(1, Col) & ""), "0.00")
    grdSub1.TextMatrix(grdSub1.Row, 3) = Format(Val(grdMinor1.TextMatrix(2, Col) & ""), "0.00")
    grdSub1.TextMatrix(grdSub1.Row, 4) = grdMinor1.TextMatrix(3, Col)
    grdSub1.TextMatrix(grdSub1.Row, 5) = grdMinor1.TextMatrix(4, Col)
End Sub
Private Sub grdMinor1_EnterCell()
    grdMinor1.ColComboList(1) = ""
    Select Case grdMinor1.Col
        Case 0
            grdMinor1.Editable = flexEDNone
            grdMinor1.Col = 1
        Case 1
            Select Case grdMinor1.Row
                Case 1, 2
                    grdMinor1.Editable = flexEDKbdMouse
                Case 3
                    grdMinor1.Editable = flexEDKbdMouse
                    grdMinor1.ColComboList(1) = "Yes|No"
                Case 4
                    grdMinor1.Editable = flexEDKbdMouse
                    grdMinor1.ColComboList(1) = "<None>|Price1|Price2|Price3|Price4|Price5|Price6"
            End Select
    End Select
End Sub
Private Sub grdMinor1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13 'Enter
        Case 45, 48 To 57, 96 To 105, 109, 110, 189
            grdMinor1.EditCell
        Case 38 'up
        Case 40 'down
    End Select
End Sub
Private Sub grdMinor1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row <> 3 Then
        If InStr(grdMinor1.EditText, ".") <> 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If
        Select Case KeyAscii
            Case 8, 13, 27, 45, 46, 48 To 57
            Case Else
                If Col = 0 Then KeyAscii = 0
        End Select
    End If
End Sub
Private Sub grdSub1_Click()
    grdMinor1.Enabled = True
    grdMinor1.TextMatrix(1, 1) = grdSub1.TextMatrix(grdSub1.Row, 2)
    grdMinor1.TextMatrix(2, 1) = grdSub1.TextMatrix(grdSub1.Row, 3)
    grdMinor1.TextMatrix(3, 1) = grdSub1.TextMatrix(grdSub1.Row, 4)
    grdMinor1.TextMatrix(4, 1) = grdSub1.TextMatrix(grdSub1.Row, 5)
End Sub
Private Sub grdSub1_RowColChange()
    grdMinor1.Enabled = True
    grdMinor1.TextMatrix(1, 1) = grdSub1.TextMatrix(grdSub1.Row, 2)
    grdMinor1.TextMatrix(2, 1) = grdSub1.TextMatrix(grdSub1.Row, 3)
    grdMinor1.TextMatrix(3, 1) = grdSub1.TextMatrix(grdSub1.Row, 4)
    grdMinor1.TextMatrix(4, 1) = grdSub1.TextMatrix(grdSub1.Row, 5)
End Sub
