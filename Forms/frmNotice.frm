VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmNotice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Notice Board"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   Icon            =   "frmNotice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Index           =   0
      Left            =   8280
      TabIndex        =   0
      Top             =   6570
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
   Begin VSFlex8Ctl.VSFlexGrid grdNotice 
      Height          =   5925
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   5355
      _cx             =   9446
      _cy             =   10451
      Appearance      =   0
      BorderStyle     =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16381166
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   380
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Index           =   1
      Left            =   6990
      TabIndex        =   2
      Top             =   6570
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
   Begin VSFlex8Ctl.VSFlexGrid grdMess 
      Height          =   6075
      Left            =   5700
      TabIndex        =   4
      Top             =   270
      Width           =   3675
      _cx             =   6482
      _cy             =   10716
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   8283198
      ForeColor       =   -2147483643
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   8283198
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      WallPaper       =   "frmNotice.frx":000C
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSForms.Image Image1 
      Height          =   6155
      Index           =   0
      Left            =   5670
      Top             =   240
      Width           =   3755
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6623;10857"
   End
   Begin MSForms.Image Image1 
      Height          =   6375
      Index           =   1
      Left            =   5580
      Top             =   120
      Width           =   3945
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "6959;11245"
   End
   Begin MSForms.Image Image3 
      Height          =   5985
      Left            =   90
      Top             =   510
      Width           =   5415
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "9551;10557"
      VariousPropertyBits=   19
   End
   Begin MSForms.Label Label1 
      Height          =   285
      Left            =   210
      TabIndex        =   3
      Top             =   180
      Width           =   3915
      Caption         =   "Details"
      Size            =   "6906;503"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image2 
      Height          =   405
      Left            =   90
      Top             =   120
      Width           =   5415
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "9551;714"
   End
End
Attribute VB_Name = "frmNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            ActiveUpdateServer "Delete from Notice_Board"
            For i = 1 To grdNotice.Rows - 1
                ActiveUpdateServer "Insert into Notice_Board (Line_no,Description,Style) Values (" & i & ",'" & grdNotice.TextMatrix(i, 1) & "', '" & grdNotice.TextMatrix(i, 2) & "')"
            Next i
            MsgBox "Noticeboard Details Updated.", vbInformation, "HeroPOS Information Message"
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    grdMess.Col = 0
    grdMess.ColWidth(0) = 180
    grdMess.TextMatrix(0, 0) = "Daily Notice Board"
    grdMess.TextMatrix(0, 1) = "Daily Notice Board"
    grdMess.MergeRow(0) = True
    grdMess.CellAlignment = flexAlignCenterCenter
    grdMess.Select 0, 0, 0, 1
    grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
    grdMess.CellFontBold = True
    
    grdNotice.ColWidth(0) = grdMess.ColWidth(0)
    grdNotice.ColWidth(1) = 3200
    grdNotice.TextMatrix(0, 1) = "Description"
    grdNotice.TextMatrix(0, 2) = "Line Style"
    grdNotice.ColAlignment(1) = flexAlignLeftCenter
    grdNotice.ColAlignment(2) = flexAlignLeftCenter
    grdMess.ColAlignment(1) = flexAlignLeftCenter
    ActiveReadServer "Select * from Notice_Board order by Line_No"
    grdNotice.Row = 0
    While Not rs.EOF
        grdNotice.Row = grdNotice.Row + 1
        grdNotice.TextMatrix(grdNotice.Row, 1) = rs.Fields("Description") & ""
        grdMess.TextMatrix(grdNotice.Row, 1) = rs.Fields("Description") & ""
        grdNotice.TextMatrix(grdNotice.Row, 2) = rs.Fields("Style") & ""
        Select Case rs.Fields("Style") & ""
            Case "Sub Header"
                grdNotice.TextMatrix(grdNotice.Row, 0) = ">"
                grdMess.TextMatrix(grdNotice.Row, 0) = ">"
                grdMess.Select grdNotice.Row, 0, grdNotice.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
            Case "Sub Header Underline"
                grdNotice.TextMatrix(grdNotice.Row, 0) = ">"
                grdMess.TextMatrix(grdNotice.Row, 0) = ">"
                grdMess.Select grdNotice.Row, 0, grdNotice.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
            Case "Normal"
                grdNotice.TextMatrix(grdNotice.Row, 0) = ""
                grdMess.TextMatrix(grdNotice.Row, 0) = ""
                grdMess.Select grdNotice.Row, 0, grdNotice.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
            Case "Normal Underline"
                grdNotice.TextMatrix(grdNotice.Row, 0) = ""
                grdMess.TextMatrix(grdNotice.Row, 0) = ""
                grdMess.Select grdNotice.Row, 0, grdNotice.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
        End Select
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub grdNotice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case grdNotice.Col
        Case 1
            grdMess.TextMatrix(grdNotice.Row, 1) = grdNotice.TextMatrix(grdNotice.Row, 1)
            If grdNotice.TextMatrix(grdNotice.Row, 2) = "" Then
                grdNotice.TextMatrix(grdNotice.Row, 2) = "Normal"
            End If
        Case 2
            Select Case grdNotice.TextMatrix(grdNotice.Row, 2)
                Case "Sub Header"
                    grdNotice.TextMatrix(grdNotice.Row, 0) = ">"
                    grdMess.TextMatrix(grdNotice.Row, 0) = ">"
                    grdMess.Select grdNotice.Row, 0, grdNotice.Row, 1
                    grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
                Case "Sub Header Underline"
                    grdNotice.TextMatrix(grdNotice.Row, 0) = ">"
                    grdMess.TextMatrix(grdNotice.Row, 0) = ">"
                    grdMess.Select grdNotice.Row, 0, grdNotice.Row, 1
                    grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
                Case "Normal"
                    grdNotice.TextMatrix(grdNotice.Row, 0) = ""
                    grdMess.TextMatrix(grdNotice.Row, 0) = ""
                    grdMess.Select grdNotice.Row, 0, grdNotice.Row, 1
                    grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
                Case "Normal Underline"
                    grdNotice.TextMatrix(grdNotice.Row, 0) = ""
                    grdMess.TextMatrix(grdNotice.Row, 0) = ""
                    grdMess.Select grdNotice.Row, 0, grdNotice.Row, 1
                    grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
            End Select
    End Select
End Sub

Private Sub grdNotice_EnterCell()
    Select Case grdNotice.Col
        Case 0
            grdNotice.Editable = flexEDNone
            grdNotice.Col = 1
        Case 1
             grdNotice.Editable = flexEDKbdMouse
        Case 2
             grdNotice.Editable = flexEDKbdMouse
             grdNotice.ColComboList(2) = "Sub Header|Sub Header Underline|Normal|Normal Underline"
    End Select
End Sub

Private Sub grdNotice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        SaveRow = grdNotice.Row
        grdNotice.RemoveItem grdNotice.Row
        grdNotice.Rows = grdNotice.Rows + 1
        grdNotice.TextMatrix(grdNotice.Rows - 1, 2) = "Normal"
        ActiveReadServer "Select * from Notice_Board order by Line_No"
        grdNotice.Row = 0
        For i = 1 To grdNotice.Rows - 1
            grdMess.TextMatrix(i, 0) = grdNotice.TextMatrix(i, 0)
            grdMess.TextMatrix(i, 1) = grdNotice.TextMatrix(i, 1)
            If grdNotice.TextMatrix(i, 1) = "" Then
                grdNotice.TextMatrix(i, 2) = "Normal"
            End If
            Select Case grdNotice.TextMatrix(i, 2)
                Case "Sub Header"
                    grdMess.TextMatrix(i, 0) = ">"
                    grdMess.Select i, 0, i, 1
                    grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
                Case "Sub Header Underline"
                    grdMess.TextMatrix(i, 0) = ">"
                    grdMess.Select i, 0, i, 1
                    grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
                Case "Normal"
                    grdMess.TextMatrix(i, 0) = ""
                    grdMess.Select i, 0, i, 1
                    grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
                Case "Normal Underline"
                    grdMess.TextMatrix(i, 0) = ""
                    grdMess.Select i, 0, i, 1
                    grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
            End Select
        Next i
        grdNotice.Row = SaveRow
    End If
End Sub

Private Sub grdNotice_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub grdNotice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
