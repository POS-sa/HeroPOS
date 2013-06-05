VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rates"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   Icon            =   "frmRates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   0
      Left            =   6180
      TabIndex        =   1
      Top             =   5310
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
      Left            =   7320
      TabIndex        =   2
      Top             =   5310
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
      Left            =   5010
      TabIndex        =   3
      Top             =   5310
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
      Height          =   5100
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   8355
      _cx             =   14737
      _cy             =   8996
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRates.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.Label lblTotRoom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2730
      TabIndex        =   6
      Top             =   5340
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   795
   End
   Begin VB.Label lblTotRate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   5340
      Width           =   1395
   End
   Begin MSForms.Image Image5 
      Height          =   315
      Index           =   16
      Left            =   960
      Top             =   5340
      Width           =   1605
      BackColor       =   16777215
      Size            =   "2831;556"
   End
   Begin MSForms.Image Image77 
      Height          =   315
      Left            =   2610
      Top             =   5340
      Visible         =   0   'False
      Width           =   2295
      BackColor       =   16777215
      Size            =   "4048;556"
   End
End
Attribute VB_Name = "frmRates"
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
            ActiveUpdateServer "Delete from Room_Rates"
            frmDetails.lblTax.Caption = ""
            For i = 1 To grdGrid.Rows - 1
                If Trim(grdGrid.TextMatrix(i, 1)) <> "" Then
                    ActiveUpdateServer "INSERT INTO [Room_Rates]([Rate_Type],[Room_Rate], [Description],[Active],[Condition]) VALUES (" & grdGrid.TextMatrix(i, 0) & "," & Val(Left(grdGrid.TextMatrix(i, 1), Len(grdGrid.TextMatrix(i, 1)) - 1)) & ",'" & Trim(grdGrid.TextMatrix(i, 2)) & "','" & Abs(Trim(grdGrid.TextMatrix(i, 3))) & "','" & Trim(grdGrid.TextMatrix(i, 4)) & "')"
                End If
                If grdGrid.TextMatrix(i, 3) = "-1" Then
                    If i = grdGrid.Rows - 1 Then
                        frmDetails.lblTax.Caption = frmDetails.lblTax.Caption & grdGrid.TextMatrix(i, 0)
                    Else
                        frmDetails.lblTax.Caption = frmDetails.lblTax.Caption & grdGrid.TextMatrix(i, 0) & "-"
                    End If
                End If
            Next i
            If frmDetails.lblTax.Caption = "" Then frmDetails.lblTax.Caption = "1"
            If Right(frmDetails.lblTax.Caption, 1) = "-" Then frmDetails.lblTax.Caption = Trim(Mid(frmDetails.lblTax.Caption, 1, Len(frmDetails.lblTax.Caption) - 1))
            saverate = frmRooms.cmbRates.Text
            frmRooms.cmbRates.Clear
            frmRooms.cmbRates.AddItem "<Select a Rate>"
            ActiveReadServer "Select * from Room_Rates where Active=1 order by Rate_Type"
            While Not rs.EOF
                i = i + 1
                frmRooms.cmbRates.AddItem rs.Fields("Rate_Type") & " - " & rs.Fields("Description") & " > " & Format(rs.Fields("Room_Rate"), "0.00")
                rs.MoveNext
            Wend
            rs.Close
            frmRooms.cmbRates.Text = "<Select a Rate>"
            For i = 0 To frmRooms.cmbRates.ListCount - 1
                If saverate = frmRooms.cmbRates.List(i) Then
                    frmRooms.cmbRates.Text = saverate
                End If
            Next i
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
    Image77.Visible = False
    lblTotRoom.Visible = False
    If grdGrid.Rows > 2 Then grdGrid.Row = 1
    If grdGrid.Rows = 1 Then
        grdGrid.Rows = 2
        grdGrid.Row = 1
        grdGrid.TextMatrix(1, 0) = 1
    End If
    grdGrid.Enabled = True
    Select Case frmRates.Tag
        Case "Check1"
            grdGrid.Enabled = False
            grdGrid.TextMatrix(0, 3) = "Active"
            cmdForms(2).Enabled = False
            ActiveReadServer "Select * from Room_Rates where Rate_Type in (" & Replace(frmCheck.lblRatestring.Caption, "-", ",") & ") order by Rate_Type"
            If rs.RecordCount > 0 Then
                grdGrid.Rows = rs.RecordCount + 1
            End If
            While Not rs.EOF
                i = i + 1
                grdGrid.TextMatrix(i, 0) = rs.Fields("Rate_Type")
                grdGrid.TextMatrix(i, 1) = Format(rs.Fields("Room_Rate"), "0.00")
                If Val(frmCheck.txt0to5.Text) <> 0 And rs.Fields("Condition") = "Children under 5" Then
                    grdGrid.TextMatrix(i, 1) = Val(frmCheck.txt0to5.Text) & " x " & Format(rs.Fields("Room_Rate"), "0.00")
                End If
                If Val(frmCheck.txt5.Text) <> 0 And rs.Fields("Condition") = "Children 5 to 12" Then
                    grdGrid.TextMatrix(i, 1) = Val(frmCheck.txt5.Text) & " x " & Format(rs.Fields("Room_Rate"), "0.00")
                End If
                If Val(frmCheck.txt12to16.Text) <> 0 And rs.Fields("Condition") = "Children 12 to 15" Then
                    grdGrid.TextMatrix(i, 1) = Val(frmCheck.txt12to16.Text) & " x " & Format(rs.Fields("Room_Rate"), "0.00")
                End If
                grdGrid.TextMatrix(i, 2) = rs.Fields("Description")
                grdGrid.TextMatrix(i, 3) = Val(rs.Fields("Active") & "")
                If rs.Fields("Condition") & "" = "" Then
                    grdGrid.TextMatrix(i, 4) = "<None>"
                Else
                    grdGrid.TextMatrix(i, 4) = rs.Fields("Condition")
                End If
                rs.MoveNext
            Wend
            rs.Close
            lblTotRate.Caption = Format(frmCheck.lblTotRate.Caption, "0.00")
        Case "Check"
            Image77.Visible = True
            lblTotRoom.Visible = True
            grdGrid.Enabled = False
            grdGrid.TextMatrix(0, 3) = "Active"
            cmdForms(2).Enabled = False
            ActiveReadServer "Select * from Room_Rates where Rate_Type in (" & Replace(frmCheckin.lblRatestring.Caption, "-", ",") & ") order by Rate_Type"
            If rs.RecordCount > 0 Then
                grdGrid.Rows = rs.RecordCount + 1
            End If
            While Not rs.EOF
                i = i + 1
                grdGrid.TextMatrix(i, 0) = rs.Fields("Rate_Type")
                grdGrid.TextMatrix(i, 1) = Format(rs.Fields("Room_Rate"), "0.00")
                If Val(frmCheckin.txt0to5.Text) <> 0 And rs.Fields("Condition") = "Children under 5" Then
                    grdGrid.TextMatrix(i, 1) = Val(frmCheckin.txt0to5.Text) & " x " & Format(rs.Fields("Room_Rate"), "0.00")
                End If
                If Val(frmCheckin.txt5.Text) <> 0 And rs.Fields("Condition") = "Children 5 to 12" Then
                    grdGrid.TextMatrix(i, 1) = Val(frmCheckin.txt5.Text) & " x " & Format(rs.Fields("Room_Rate"), "0.00")
                End If
                If Val(frmCheckin.txt12to16.Text) <> 0 And rs.Fields("Condition") = "Children 12 to 15" Then
                    grdGrid.TextMatrix(i, 1) = Val(frmCheckin.txt12to16.Text) & " x " & Format(rs.Fields("Room_Rate"), "0.00")
                End If
                grdGrid.TextMatrix(i, 2) = rs.Fields("Description")
                grdGrid.TextMatrix(i, 3) = Val(rs.Fields("Active") & "")
                If rs.Fields("Condition") & "" = "" Then
                    grdGrid.TextMatrix(i, 4) = "<None>"
                Else
                    grdGrid.TextMatrix(i, 4) = rs.Fields("Condition")
                End If
                rs.MoveNext
            Wend
            rs.Close
            lblTotRate.Caption = Format(frmCheckin.lblTotRate.Caption, "0.00")
            lblTotRoom.Caption = Val(frmCheckin.txtNights.Text) - Val(frmCheckin.txtFree.Text) & " x Nights = " & Format(Val(frmCheckin.lblTotRate.Caption) * (Val(frmCheckin.txtNights.Text) - Val(frmCheckin.txtFree.Text)), "0.00")
        Case Else
            i = 0
            ActiveReadServer "Select * from Room_Rates order by Rate_Type"
            If rs.RecordCount > 0 Then
                grdGrid.Rows = rs.RecordCount + 1
            End If
            While Not rs.EOF
                i = i + 1
                grdGrid.TextMatrix(i, 0) = rs.Fields("Rate_Type")
                grdGrid.TextMatrix(i, 1) = Format(rs.Fields("Room_Rate"), "0.00")
                grdGrid.TextMatrix(i, 2) = rs.Fields("Description")
                grdGrid.TextMatrix(i, 3) = Val(rs.Fields("Active") & "")
                If rs.Fields("Condition") & "" = "" Then
                    grdGrid.TextMatrix(i, 4) = "<None>"
                Else
                    grdGrid.TextMatrix(i, 4) = rs.Fields("Condition")
                End If
                rs.MoveNext
            Wend
            rs.Close
    End Select
End Sub
Private Sub Form_Load()
    grdGrid.ColWidth(0) = grdGrid.Width * 0.07
    grdGrid.ColWidth(1) = grdGrid.Width * 0.15
    grdGrid.ColWidth(2) = grdGrid.Width * 0.3
    grdGrid.ColAlignment(4) = flexAlignLeftCenter
    grdGrid.ColAlignment(1) = flexAlignLeftCenter
    grdGrid.TextMatrix(0, 4) = "Conditions"
End Sub
Private Sub grdGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        If Right(grdGrid.Text, 1) <> "%" And Trim(grdGrid.Text) <> "" Then
            grdGrid.Text = Format(grdGrid.Text, "0.00")
        End If
    End If
    If Col = 4 Then
        If grdGrid.Text = "" Then grdGrid.Text = "<None>"
    End If
End Sub
Private Sub grdGrid_EnterCell()
    If frmRates.Tag = "Check" Then
        grdGrid.ColComboList(4) = ""
        Exit Sub
    End If
    If grdGrid.Col = 4 Then
        grdGrid.ColComboList(4) = "<None>|Per Suite|Adult Single|Adult Double|Three or More Adults|Children under 5|Children 5 to 12|Children 12 to 15|Children 16 and Over|Representatives|Pensioner Single|Pensioner Double"
    End If
End Sub
Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
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
                    grdGrid.TextMatrix(grdGrid.Row, 4) = "<None>"
                End If
            End If
    End Select
End Sub
Private Sub grdGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = 0
    End Select
    If Col = 1 Then
        Select Case KeyAscii
        Case 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
    End If
End Sub


