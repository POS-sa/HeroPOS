VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHolidays 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Public Holidays..."
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
   Icon            =   "frmHolidays.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   6360
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4425
      _cx             =   7805
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
      SelectionMode   =   0
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
      FormatString    =   $"frmHolidays.frx":000C
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
      Begin MSComCtl2.MonthView DTPicker 
         Height          =   2760
         Left            =   1920
         TabIndex        =   1
         Top             =   300
         Visible         =   0   'False
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   4868
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16239822
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowToday       =   0   'False
         StartOfWeek     =   50135041
         TitleBackColor  =   16239822
         CurrentDate     =   39226
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   2
         Top             =   11700
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Left            =   3330
      TabIndex        =   3
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
      Left            =   2040
      TabIndex        =   4
      Top             =   6510
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Save"
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
Attribute VB_Name = "frmHolidays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
    Unload Me
End Sub
Private Sub cmdOk1_Click()
    ActiveUpdateServer "Delete from Holidays"
    For i = 1 To grdMain.Rows - 1
        If Trim(grdMain.TextMatrix(i, 0)) <> "" Then
            If Trim(grdMain.TextMatrix(i, 1)) <> "" Then
                ActiveUpdateServer "Insert into Holidays (Description,Date_Time) values ('" & Trim(grdMain.TextMatrix(i, 0)) & "','" & Format(grdMain.TextMatrix(i, 2), "YYYY-MM-DD") & "')"
            End If
        End If
    Next i
    Unload Me
End Sub
Private Sub DTPicker_DateClick(ByVal DateClicked As Date)
    grdMain.TextMatrix(grdMain.Row, grdMain.Col) = Format(DTPicker.Value, "DDDD DD MMMM YYYY")
    grdMain.TextMatrix(grdMain.Row, 2) = DTPicker.Value
    DTPicker.Visible = False
    grdMain.SetFocus
End Sub

Private Sub Form_Load()
    grdMain.ColHidden(2) = True
    grdMain.Rows = 1
    grdMain.TextMatrix(0, 0) = "Description"
    grdMain.TextMatrix(0, 1) = "Day"
    grdMain.ColWidth(0) = grdMain.Width - DTPicker.Width - 30
    grdMain.ColDataType(2) = flexDTDate
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    ActiveReadServer "Select * from Holidays order by Date_Time"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Description")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = Format(rs.Fields("Date_Time"), "DDDD DD MMMM YYYY")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = rs.Fields("Date_Time")
        rs.MoveNext
    Wend
    rs.Close
    If grdMain.Rows = 1 Then
        grdMain.Rows = 2
    End If
    grdMain.Row = 1
End Sub
Private Sub grdMain_DblClick()
    On Error Resume Next
    Select Case grdMain.Col
        Case 0
            grdMain.EditCell
        Case 1
            If grdMain.RowPos(grdMain.Row) + DTPicker.Height > frmSpecials.Height Then
                DTPicker.top = grdMain.RowPos(grdMain.Row) - DTPicker.Height + grdMain.RowHeight(grdMain.Row)
                DTPicker.Left = grdMain.ColPos(grdMain.Col)
            Else
                DTPicker.top = grdMain.RowPos(grdMain.Row)
                DTPicker.Left = grdMain.ColPos(grdMain.Col)
            End If
            DTPicker.Visible = True
            DTPicker.Value = Format(grdMain.TextMatrix(grdMain.Row, 1), "YYYY-MM-DD")
    End Select
    On Error GoTo 0
End Sub
Private Sub grdMain_GotFocus()
    DTPicker.Visible = False
End Sub
Private Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            Select Case grdMain.Col
                Case 0
                    grdMain.EditCell
                Case 1
                    If grdMain.RowPos(grdMain.Row) + DTPicker.Height > frmSpecials.Height Then
                        DTPicker.top = grdMain.RowPos(grdMain.Row) - DTPicker.Height + grdMain.RowHeight(grdMain.Row)
                        DTPicker.Left = grdMain.ColPos(grdMain.Col)
                    Else
                        DTPicker.top = grdMain.RowPos(grdMain.Row)
                        DTPicker.Left = grdMain.ColPos(grdMain.Col)
                    End If
                    DTPicker.Visible = True
                    DTPicker.Value = Format(grdMain.TextMatrix(grdMain.Row, 1), "YYYY-MM-DD")
            End Select
        Case 40
            If grdMain.Row = grdMain.Rows - 1 And grdMain.TextMatrix(grdMain.Row, 0) <> "" Then
                grdMain.Rows = grdMain.Rows + 1
                grdMain.Row = grdMain.Rows - 1
                grdMain.Col = 0
                KeyCode = 0
                On Error GoTo 0
                Exit Sub
            End If
            If grdMain.TextMatrix(grdMain.Row, 0) = "" Then
                grdMain.Row = 1
                grdMain.ShowCell 1, 1
                KeyCode = 0
                On Error GoTo 0
                Exit Sub
            End If
            If grdMain.Row = grdMain.Rows - 2 And grdMain.TextMatrix(grdMain.Rows - 1, 0) = "" Then
                grdMain.Col = 0
            End If
        Case 46
            grdMain.RemoveItem (grdMain.Row)
        Case Else
            If grdMain.Col = 0 Then grdMain.EditCell
    End Select
    On Error GoTo 0
End Sub
Private Sub grdMain_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub grdMain_RowColChange()
    DTPicker.Visible = False
End Sub
