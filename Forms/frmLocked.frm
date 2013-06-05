VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLocked 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unlock Tabs"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   Icon            =   "frmLocked.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   4980
      TabIndex        =   0
      ToolTipText     =   " Click to Search.... "
      Top             =   4350
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
      Left            =   3690
      TabIndex        =   1
      ToolTipText     =   " Click to Search.... "
      Top             =   4350
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Unlock"
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
   Begin VSFlex8Ctl.VSFlexGrid grdList 
      Height          =   4200
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6145
      _cx             =   10839
      _cy             =   7408
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
      FormatString    =   $"frmLocked.frx":000C
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
Attribute VB_Name = "frmLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    For i = 1 To grdList.Rows - 1
        If grdList.ValueMatrix(i, 2) <> 0 Then
            Select Case Me.Tag
                Case "Tables"
                    ActiveUpdateServer "Update Table_Listing set Locked = 0 where Table_No = " & grdList.TextMatrix(i, 0)
                Case "Tabs"
                    ActiveUpdateServer "Update Tab_Listing set Locked = 0 where Tab_No = " & Val(Trim(Mid(grdList.TextMatrix(i, 0), 1, InStr(grdList.TextMatrix(i, 0), "-") - 1)))
            End Select
        End If
    Next i
    Select Case Me.Tag
        Case "Tables"
            MsgBox "Selected Tables were Unlocked", vbInformation, "HeroPOS"
            grdList.Rows = 1
            ActiveReadServer "Select * from Table_Listing_View where Locked= 1 order by Table_No"
            While Not rs.EOF
                grdList.Rows = grdList.Rows + 1
                grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Table_No")
                grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                grdList.TextMatrix(grdList.Rows - 1, 2) = "0"
                rs.MoveNext
            Wend
            rs.Close
        Case "Tabs"
            MsgBox "Selected Tabs were Unlocked", vbInformation, "HeroPOS"
            grdList.Rows = 1
            ActiveReadServer "Select * from Tab_Listing_View where Locked= 1 order by Tab_No"
            While Not rs.EOF
                grdList.Rows = grdList.Rows + 1
                grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Tab_No") & " - " & rs.Fields("Tab_Name")
                grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                grdList.TextMatrix(grdList.Rows - 1, 2) = "0"
                rs.MoveNext
            Wend
            rs.Close
    End Select
    On Error GoTo 0
End Sub
Private Sub Form_Activate()
    Select Case Me.Tag
        Case "Tables"
            Me.Caption = "Unlock Selected Tables..."
            grdList.TextMatrix(0, 0) = "Table Number"
            grdList.TextMatrix(0, 1) = "User"
            grdList.TextMatrix(0, 2) = "Select"
            grdList.ColWidth(0) = grdList.Width * 0.25
            grdList.ColWidth(1) = grdList.Width * 0.6
            grdList.ColAlignment(0) = flexAlignLeftCenter
            grdList.ColAlignment(1) = flexAlignLeftCenter
            grdList.ColAlignment(2) = flexAlignCenterCenter
            grdList.ColDataType(2) = flexDTBoolean
            grdList.Rows = 1
            ActiveReadServer "Select * from Table_Listing_View where Locked= 1 order by Table_No"
            While Not rs.EOF
                grdList.Rows = grdList.Rows + 1
                grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Table_No")
                grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                grdList.TextMatrix(grdList.Rows - 1, 2) = "0"
                rs.MoveNext
            Wend
            rs.Close
        Case "Tabs"
            Me.Caption = "Unlock Selected Tabs..."
            grdList.TextMatrix(0, 0) = "Tab Description"
            grdList.TextMatrix(0, 1) = "User"
            grdList.TextMatrix(0, 2) = "Select"
            grdList.ColWidth(0) = grdList.Width * 0.35
            grdList.ColWidth(1) = grdList.Width * 0.5
            grdList.ColAlignment(0) = flexAlignLeftCenter
            grdList.ColAlignment(1) = flexAlignLeftCenter
            grdList.ColAlignment(2) = flexAlignCenterCenter
            grdList.ColDataType(2) = flexDTBoolean
            grdList.Rows = 1
            ActiveReadServer "Select * from Tab_Listing_View where Locked= 1 order by Tab_No"
            While Not rs.EOF
                grdList.Rows = grdList.Rows + 1
                grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Tab_No") & " - " & rs.Fields("Tab_Name")
                grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                grdList.TextMatrix(grdList.Rows - 1, 2) = "0"
                rs.MoveNext
            Wend
            rs.Close
    End Select
End Sub
Private Sub grdList_EnterCell()
    Select Case grdList.Col
        Case 2
            grdList.Editable = flexEDKbdMouse
        Case Else
            grdList.Editable = flexEDNone
    End Select
End Sub
