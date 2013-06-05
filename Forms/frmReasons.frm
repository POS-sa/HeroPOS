VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmReasons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Overide Reasons"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmReasons.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   0
      Left            =   2340
      TabIndex        =   0
      Top             =   6270
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   6270
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Help"
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
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   2
      Left            =   1170
      TabIndex        =   2
      Top             =   6270
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "&Ok"
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
      Height          =   6120
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   4545
      _cx             =   8017
      _cy             =   10795
      Appearance      =   1
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
      FocusRect       =   1
      HighLight       =   1
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
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmReasons.frx":000C
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
End
Attribute VB_Name = "frmReasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForms_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
        Case 2
            ActiveUpdateServer "Delete from Reasons"
            For i = 1 To grdGrid.Rows - 1
                If Trim(grdGrid.TextMatrix(i, 1)) <> "" Then
                    ActiveUpdateServer "INSERT INTO [Reasons]([Reason_No], [Reason_Name]) VALUES (" & grdGrid.TextMatrix(i, 0) & ",'" & Trim(grdGrid.TextMatrix(i, 1)) & "')"
                End If
            Next i
            MsgBox "Overide reasons Updated.", vbInformation, "HeroPOS Information Message"
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
    If grdGrid.Rows > 1 Then
        grdGrid.Row = 1
        grdGrid.TextMatrix(1, 0) = "1"
    End If
End Sub
Private Sub Form_Load()
    grdGrid.TextMatrix(0, 1) = "Reason Description"
    grdGrid.Rows = 1
    ActiveReadServer "Select * from Reasons order by Reason_No"
    While Not rs.EOF
        grdGrid.Rows = grdGrid.Rows + 1
        grdGrid.Row = grdGrid.Rows - 1
        grdGrid.TextMatrix(grdGrid.Rows - 1, 0) = rs.Fields("Reason_No")
        grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = rs.Fields("Reason_Name")
        rs.MoveNext
    Wend
    rs.Close
    If grdGrid.Rows = 1 Then grdGrid.Rows = 2
End Sub
Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46
            grdGrid.RemoveItem grdGrid.Rows - 1
        Case 38
            If grdGrid.Row = grdGrid.Rows - 1 Then
                If grdGrid.Text = "" Then
                    grdGrid.RemoveItem grdGrid.Rows - 1
                End If
            End If
        Case 40
            If grdGrid.Row = grdGrid.Rows - 1 Then
                If grdGrid.Text <> "" Then
                    grdGrid.Rows = grdGrid.Rows + 1
                    grdGrid.Row = grdGrid.Rows - 1
                    grdGrid.TextMatrix(grdGrid.Row, 0) = Val(grdGrid.TextMatrix(grdGrid.Row - 1, 0)) + 1
                End If
            End If
        Case 13, 45, 48 To 57, 65 To 90, 96 To 105, 109, 110, 189
            grdGrid.EditCell
    End Select
End Sub
Private Sub grdGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = 0
    End Select
End Sub

