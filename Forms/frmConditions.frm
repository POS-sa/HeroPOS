VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmConditions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Condition of Stay"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   Icon            =   "frmConditions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid grdGrid 
      Height          =   3750
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9765
      _cx             =   17224
      _cy             =   6615
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
      Rows            =   25
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmConditions.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   345
      Left            =   8670
      TabIndex        =   1
      Top             =   3960
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
      Left            =   7440
      TabIndex        =   2
      Top             =   3960
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
   Begin MSForms.Image Image2 
      Height          =   3765
      Left            =   90
      Top             =   90
      Width           =   9795
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "17277;6641"
   End
End
Attribute VB_Name = "frmConditions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ActiveUpdateServer "Delete from Conditions"
    DoEvents
    For i = 1 To grdGrid.Rows - 1
        ActiveUpdateServer "Insert into Conditions (Description,Bulleted) values ('" & grdGrid.TextMatrix(i, 1) & "'," & grdGrid.ValueMatrix(i, 2) & ")"
        DoEvents
    Next i
    Unload Me
End Sub

Private Sub Form_Load()
    grdGrid.RowHeight(0) = 580
    grdGrid.Cols = 3
    grdGrid.TextMatrix(0, 0) = "Bullet"
    grdGrid.TextMatrix(0, 1) = "Description"
    grdGrid.TextMatrix(0, 2) = "Bulleted"
    grdGrid.ColWidth(0) = grdGrid.Width * 0.05
    grdGrid.ColWidth(1) = grdGrid.Width * 0.8
    grdGrid.ColWidth(2) = grdGrid.Width * 0.15
    grdGrid.ColDataType(2) = flexDTBoolean
    grdGrid.ColAlignment(0) = flexAlignCenterCenter
    grdGrid.ColAlignment(1) = flexAlignLeftCenter
    grdGrid.ColAlignment(2) = flexAlignCenterCenter
    grdGrid.Row = 0
    ActiveReadServer "Select * from Conditions"
    While Not rs.EOF
        grdGrid.Row = grdGrid.Row + 1
        grdGrid.TextMatrix(grdGrid.Row, 1) = rs.Fields("Description")
        grdGrid.TextMatrix(grdGrid.Row, 2) = rs.Fields("Bulleted")
        If rs.Fields("Bulleted") <> 0 Then
            grdGrid.TextMatrix(grdGrid.Row, 0) = ">"
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub grdGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    DoEvents
    If grdGrid.Col = 2 Then
        Select Case grdGrid.ValueMatrix(Row, 2)
            Case True
                grdGrid.TextMatrix(Row, 0) = ">"
            Case False
                grdGrid.TextMatrix(Row, 0) = ""
        End Select
    End If
End Sub

Private Sub grdGrid_EnterCell()
    Select Case grdGrid.Col
        Case 1
            grdGrid.Editable = flexEDKbdMouse
        Case 2
            grdGrid.Editable = flexEDKbdMouse
    End Select
End Sub
Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13 'Enter
        Case 46
            grdGrid.TextMatrix(grdGrid.Row, grdGrid.Col) = ""
        Case 45, 48 To 57, 65 To 90, 96 To 105, 109, 110, 189
            grdGrid.EditCell
        Case Else
    End Select
End Sub
Private Sub grdGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
