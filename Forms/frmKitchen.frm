VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmKitchen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kitchen Printer Setup"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   Icon            =   "frmKitchen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDepartments 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   30
      ScaleHeight     =   3825
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   10275
      Begin VSFlex8Ctl.VSFlexGrid grdMinor1 
         Height          =   3810
         Left            =   5340
         TabIndex        =   1
         Top             =   0
         Width           =   4065
         _cx             =   7170
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
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   1500
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmKitchen.frx":000C
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   700
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmKitchen.frx":0084
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
         FormatString    =   $"frmKitchen.frx":00FC
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
      Left            =   8220
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
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
      Left            =   6990
      TabIndex        =   5
      Top             =   3960
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Set"
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
Attribute VB_Name = "frmKitchen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ActiveUpdateServer "Update Products set Kitchen1='" & grdMinor1.TextMatrix(1, 1) & "' where Department_No = '" & grdSub1.TextMatrix(grdSub1.Row, 0) & "'"
    DoEvents
    ActiveUpdateServer "Update Products set Kitchen2='" & grdMinor1.TextMatrix(2, 1) & "' where Department_No = '" & grdSub1.TextMatrix(grdSub1.Row, 0) & "'"
    DoEvents
    MsgBox "Kitchen Printers Setup Changed", vbInformation, "HeroPOS"
    DoEvents
End Sub
Private Sub Form_Load()
    grdMinor1.Cols = 2
    grdSub1.ColHidden(2) = True
    grdSub1.ColHidden(3) = True
    grdSub1.ColHidden(4) = True
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
    grdMinor1.ColAlignment(1) = flexAlignLeftCenter
    grdMajor1.Rows = 1
    grdSub1.Rows = 1
    grdMinor1.Rows = 3
    
    grdMinor1.TextMatrix(1, 0) = "Kitchen Printer 1"
    grdMinor1.TextMatrix(2, 0) = "Kitchen Printer 2"
    grdMinor1.TextMatrix(1, 1) = "<None>"
    grdMinor1.TextMatrix(2, 1) = "<None>"
    
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
Private Sub grdMajor1_Click()
    cmdOk.Enabled = False
    grdSub1.Rows = 1
    grdMinor1.Rows = 3
    ActiveReadServer1 "Select * From Departments where Dept_Type=1 and Dept_Parent='" & grdMajor1.TextMatrix(grdMajor1.Row, 0) & "'"
    i = 0
    While Not rs1.EOF
        i = i + 1
        grdSub1.Rows = grdSub1.Rows + 1
        grdSub1.TextMatrix(i, 0) = rs1.Fields("Department_No")
        grdSub1.TextMatrix(i, 1) = rs1.Fields("Dept_Name")
        rs1.MoveNext
    Wend
End Sub
Private Sub grdMinor1_EnterCell()
    Dim x As Printer
    grdMinor1.ColComboList(1) = ""
    grdMinor1.ColComboList(1) = "<None>|"
    For Each x In Printers
        grdMinor1.ColComboList(1) = grdMinor1.ColComboList(1) & "|" & x.DeviceName
    Next
End Sub

Private Sub grdSub1_Click()
    grdMinor1.Enabled = True
    cmdOk.Enabled = True
End Sub
