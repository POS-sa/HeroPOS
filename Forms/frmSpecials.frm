VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSpecials 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Specials...."
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00D6DAFE&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   4500
      ScaleHeight     =   315
      ScaleWidth      =   3015
      TabIndex        =   13
      Top             =   6990
      Width           =   3045
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inactive Specials"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   810
         TabIndex        =   15
         Top             =   30
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEFDD9&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1380
      ScaleHeight     =   315
      ScaleWidth      =   3015
      TabIndex        =   12
      Top             =   6990
      Width           =   3045
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Active Specials"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   810
         TabIndex        =   14
         Top             =   30
         Width           =   2325
      End
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   90
      ScaleHeight     =   2925
      ScaleWidth      =   5355
      TabIndex        =   9
      Top             =   450
      Visible         =   0   'False
      Width           =   5355
      Begin btButtonEx.ButtonEx cmdOk 
         Height          =   315
         Left            =   4110
         TabIndex        =   10
         Top             =   2460
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
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
      Begin MSComCtl2.MonthView mthView 
         Height          =   2310
         Left            =   90
         TabIndex        =   11
         Top             =   90
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16239822
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxSelCount     =   365
         MonthColumns    =   2
         MonthBackColor  =   16777215
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         StartOfWeek     =   73203714
         TitleBackColor  =   16761281
         TrailingForeColor=   -2147483639
         CurrentDate     =   38701
      End
      Begin MSForms.Image Image5 
         Height          =   2925
         Left            =   0
         Top             =   0
         Width           =   5355
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "9446;5159"
      End
      Begin MSForms.Image Image6 
         Height          =   2805
         Left            =   60
         Top             =   60
         Width           =   5235
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "9234;4948"
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   6360
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Width           =   10035
      _cx             =   17701
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSpecials.frx":0000
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
      Begin MSComCtl2.MonthView DTPicker 
         Height          =   2760
         Left            =   2910
         TabIndex        =   16
         Top             =   3360
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
         StartOfWeek     =   73203713
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
         TabIndex        =   1
         Top             =   11700
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   375
      Left            =   2490
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   8421504
      Caption         =   "¦"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Left            =   8910
      TabIndex        =   6
      Top             =   6990
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
      Left            =   7620
      TabIndex        =   7
      Top             =   6990
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
   Begin btButtonEx.ButtonEx cmdDelete 
      Height          =   345
      Left            =   90
      TabIndex        =   8
      Top             =   6990
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Delete"
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
   Begin MSForms.Image Image1 
      Height          =   6405
      Left            =   90
      Top             =   510
      Width           =   10065
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "17754;11298"
   End
   Begin MSForms.ComboBox cmb1 
      Height          =   375
      Left            =   6810
      TabIndex        =   5
      Top             =   90
      Width           =   3345
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "5900;661"
      ListRows        =   10
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmb2 
      Height          =   375
      Left            =   2970
      TabIndex        =   4
      Top             =   90
      Width           =   3795
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "6694;661"
      ListRows        =   10
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblDate 
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   180
      Width           =   2235
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "1 Feb 2006 to 13 Feb 2006"
      Size            =   "3942;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image2 
      Height          =   375
      Left            =   90
      Top             =   90
      Width           =   2385
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "4207;661"
   End
End
Attribute VB_Name = "frmSpecials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonEx1_Click()
    dtPicker.Visible = False
    Select Case ButtonEx1.Value
        Case 0
            picDate.Visible = True
        Case 1
            picDate.Visible = False
            If picDate.Visible = False Then Selection_Change
    End Select
End Sub
Private Sub cmb1_Change()
   If cmb1.Tag = "" Then
        Selection_Change
    End If
End Sub
Private Sub cmb1_GotFocus()
    picDate.Visible = False
    dtPicker.Visible = False
End Sub
Private Sub cmb2_Change()
    If cmb2.Tag = "" Then
        Selection_Change
    End If
End Sub
Private Sub cmb2_GotFocus()
    picDate.Visible = False
    dtPicker.Visible = False
End Sub
Private Sub cmdDelete_Click()
    ActiveUpdateServer "Delete from Specials where Line_No = " & grdMain.TextMatrix(grdMain.Row, 6)
    DoEvents
    grdMain.RemoveItem grdMain.Row
    If grdMain.Rows = 1 Then
        grdMain.Rows = grdMain.Rows + 1
        grdMain.Row = grdMain.Rows - 1
    End If
    picDate.Visible = False
    dtPicker.Visible = False
End Sub
Private Sub cmdEnd_Click()
    dtPicker.Visible = False
    Unload Me
End Sub
Private Sub cmdOk_Click()
    picDate.Visible = False
    If picDate.Visible = False Then Selection_Change
End Sub
Private Sub Selection_Change()
    
    If cmb2.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(cmb2.Text, 1, InStrRev(cmb2.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(cmb2.Text, 1, InStrRev(cmb2.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(cmb2.Text, 1, InStrRev(cmb2.Text, "-") - 2)
        End If
    End If
    grdMain.Rows = 1
    Select Case cmb1.Text
        Case "Active Specials"
            ActiveReadServer "Select * from Specials_View where Active =1 and Product_Code in (Select Product_Code from Products where Department_No like '" & DeptString & "') order by Description"
        Case "Inactive Specials"
            ActiveReadServer "Select * from Specials_View where Active =0 and Product_Code in (Select Product_Code from Products where Department_No like '" & DeptString & "') order by Description"
        Case Else
            ActiveReadServer "Select * from Specials_View where Product_Code in (Select Product_Code from Products where Department_No like '" & DeptString & "') order by Description"
    End Select
    If rs.RecordCount > 1 Then grdMain.Rows = 1
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Selling_Price"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("StartDate"), "YYYY-MM-DD")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("StopDate"), "YYYY-MM-DD")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Price"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = rs.Fields("Line_No")
        If rs.Fields("Active") = 1 Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 5) = &HCEFDD9
        End If
        If rs.Fields("Active") = 0 Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 5) = &HD6DAFE
        End If
        rs.MoveNext
    Wend
    grdMain.Row = grdMain.Rows - 1
    grdMain.ShowCell grdMain.Rows - 1, 0
End Sub
Private Sub cmdOk1_Click()
    DoEvents
    For i = 1 To grdMain.Rows - 1
        If grdMain.TextMatrix(i, 0) <> "" Then
            ActiveUpdateServer "Delete from Specials where Line_No = " & grdMain.TextMatrix(i, 6)
            DoEvents
            If Format(grdMain.TextMatrix(i, 4), "YYYY-MM-DD") < Format(Date, "YYYY-MM-DD") Then
                ActiveUpdateServer "INSERT INTO [Specials]([Product_Code], [StartDate], [StopDate], [Price], [Active]) " & _
                "VALUES('" & grdMain.TextMatrix(i, 0) & "', '" & grdMain.TextMatrix(i, 3) & "', '" & grdMain.TextMatrix(i, 4) & "', '" & grdMain.TextMatrix(i, 5) & "', '0')"
            Else
            ActiveUpdateServer "INSERT INTO [Specials]([Product_Code], [StartDate], [StopDate], [Price], [Active]) " & _
            "VALUES('" & grdMain.TextMatrix(i, 0) & "', '" & grdMain.TextMatrix(i, 3) & "', '" & grdMain.TextMatrix(i, 4) & "', '" & grdMain.TextMatrix(i, 5) & "', '1')"
            End If
        End If
    Next i
    dtPicker.Visible = False
    Unload Me
End Sub

Private Sub DTPicker_DateClick(ByVal DateClicked As Date)
    grdMain.TextMatrix(grdMain.Row, grdMain.Col) = Format(dtPicker.Value, "YYYY-MM-DD")
    dtPicker.Visible = False
    If grdMain.Col = 4 Then
        If Format(grdMain.TextMatrix(grdMain.Row, 4), "YYYY-MM-DD") < Format(grdMain.TextMatrix(grdMain.Row, 3), "YYYY-MM-DD") Then
            grdMain.TextMatrix(grdMain.Row, 4) = grdMain.TextMatrix(grdMain.Row, 3)
        End If
    End If
    If Format(grdMain.TextMatrix(grdMain.Row, 4), "YYYY-MM-DD") < Format(Date, "YYYY-MM-DD") Then
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 5) = &HD6DAFE
    Else
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 5) = &HCEFDD9
    End If
    grdMain.SetFocus
End Sub
Private Sub Form_Activate()
    If frmSpecials.Tag = "1" Then
        frmSpecials.Tag = ""
        Exit Sub
    End If
    grdMain.SetFocus
    grdMain.Col = 0
    If grdMain.Rows = 1 Then
        grdMain.Rows = 2
        grdMain.Row = grdMain.Row + 1
    End If
End Sub
Private Sub Form_Click()
    picDate.Visible = False
End Sub
Private Sub Form_Load()
    lblDate.Caption = Format(Date, "DD MMM YYYY") & " to " & Format(Date, "DD MMM YYYY")
    mthView.Value = Date
    mthView.SelStart = Date
    mthView.SelEnd = Date
    picDate.Visible = False
    cmb2.Tag = "1"
    cmb2.Clear
    ActiveReadServer "Select Department_No,Dept_Name from Departments order by Department_no"
    cmb2.AddItem "<All Departments>"
    While Not rs.EOF
        cmb2.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmb2.Text = "<All Departments>"
    cmb2.Tag = ""
    cmb1.Tag = "1"
    cmb1.Clear
    cmb1.AddItem "<All Specials>"
    cmb1.AddItem "Active Specials"
    cmb1.AddItem "Inactive Specials"
    cmb1.Text = "Active Specials"
    cmb1.Tag = ""
    grdMain.RowHeight(0) = 450
    grdMain.Cols = 7
    grdMain.TextMatrix(0, 0) = "Product Code"
    grdMain.TextMatrix(0, 1) = "Description"
    grdMain.TextMatrix(0, 2) = "Selling Price"
    grdMain.TextMatrix(0, 3) = "Start Date"
    grdMain.TextMatrix(0, 4) = "Stop Date"
    grdMain.TextMatrix(0, 5) = "Special Price"
    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.31
    grdMain.ColWidth(2) = grdMain.Width * 0.13
    grdMain.ColWidth(3) = grdMain.Width * 0.13
    grdMain.ColWidth(4) = grdMain.Width * 0.13
    grdMain.ColWidth(5) = grdMain.Width * 0.13
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignRightCenter
    grdMain.ColAlignment(3) = flexAlignLeftCenter
    grdMain.ColAlignment(4) = flexAlignLeftCenter
    grdMain.ColAlignment(5) = flexAlignRightCenter
    grdMain.ColHidden(6) = True
    grdMain.Rows = 1
    ActiveReadServer "Select * from Specials_View where " & _
    " Active =1 order by Description"
    If rs.RecordCount > 1 Then grdMain.Rows = 1
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Description")
        grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Selling_Price"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = Format(rs.Fields("StartDate"), "YYYY-MM-DD")
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = Format(rs.Fields("StopDate"), "YYYY-MM-DD")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = Format(rs.Fields("Price"), "0.00")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = rs.Fields("Line_No")
        If rs.Fields("Active") = 1 Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 5) = &HCEFDD9
        End If
        If rs.Fields("Active") = 0 Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 5) = &HD6DAFE
        End If
        rs.MoveNext
    Wend
    If rs.RecordCount > 0 Then
        grdMain.Row = grdMain.Rows - 1
        grdMain.ShowCell grdMain.Rows - 1, 0
    End If
    rs.Close
End Sub
Private Sub grdMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    grdMain.TextMatrix(grdMain.Row, 5) = Format(grdMain.TextMatrix(grdMain.Row, 5), "0.00")
End Sub
Private Sub grdMain_BeforeSort(ByVal Col As Long, Order As Integer)
    If grdMain.TextMatrix(grdMain.Rows - 1, 0) = "" Then
        grdMain.RemoveItem grdMain.Rows - 1
    End If
End Sub
Private Sub grdMain_DblClick()
    On Error Resume Next
    Select Case grdMain.Col
        Case 3, 4
            If grdMain.RowPos(grdMain.Row) + dtPicker.Height > frmSpecials.Height Then
                dtPicker.top = grdMain.RowPos(grdMain.Row) - dtPicker.Height + grdMain.RowHeight(grdMain.Row)
                dtPicker.Left = grdMain.ColPos(grdMain.Col)
            Else
                dtPicker.top = grdMain.RowPos(grdMain.Row)
                dtPicker.Left = grdMain.ColPos(grdMain.Col)
            End If
            dtPicker.Visible = True
            dtPicker.Value = Format(grdMain.TextMatrix(grdMain.Row, grdMain.Col), "YYYY-MM-DD")
    End Select
    On Error GoTo 0
End Sub
Private Sub grdMain_GotFocus()
    picDate.Visible = False
    dtPicker.Visible = False
End Sub
Private Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 13
            Select Case grdMain.Col
                Case 0, 1
                    frmSpecials.Tag = "1"
                    Load frmSearch
                    frmSearch.Tag = "Specials"
                    frmSearch.Show vbModal
                    Select Case frmSearch.Tag
                        Case ""
                        Case Else
                            ActiveReadServer "Select * from Products where Product_code ='" & Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1)) & "'"
                            If rs.RecordCount > 0 Then
                                grdMain.TextMatrix(grdMain.Row, 0) = rs.Fields("Product_Code")
                                grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Description")
                                grdMain.TextMatrix(grdMain.Row, 3) = Format(Date, "YYYY-MM-DD")
                                grdMain.TextMatrix(grdMain.Row, 4) = Format(Date, "YYYY-MM-DD")
                                grdMain.TextMatrix(grdMain.Row, 2) = Format(rs.Fields("Selling_Price"), "0.00")
                                grdMain.TextMatrix(grdMain.Row, 5) = "0.00"
                            End If
                            rs.Close
                            frmSearch.Tag = ""
                    End Select
                    Unload frmSearch
                    grdMain.Col = 3
                Case 3, 4
                    If grdMain.RowPos(grdMain.Row) + dtPicker.Height > frmSpecials.Height Then
                        dtPicker.top = grdMain.RowPos(grdMain.Row) - dtPicker.Height + grdMain.RowHeight(grdMain.Row)
                        dtPicker.Left = grdMain.ColPos(grdMain.Col)
                    Else
                        dtPicker.top = grdMain.RowPos(grdMain.Row)
                        dtPicker.Left = grdMain.ColPos(grdMain.Col)
                    End If
                    dtPicker.Visible = True
                    dtPicker.Value = Format(grdMain.TextMatrix(grdMain.Row, grdMain.Col), "YYYY-MM-DD")
                Case 5
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
        Case Else
            If grdMain.Col = 5 Then grdMain.EditCell
    End Select
    On Error GoTo 0
End Sub
Private Sub grdMain_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(grdMain.EditText, ".") <> 0 And KeyAscii = 46 Then
        KeyAscii = 0
    End If
    Select Case KeyAscii
        Case 8, 13, 27, 45, 46, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub grdMain_RowColChange()
    dtPicker.Visible = False
End Sub

Private Sub mthView_LostFocus()
    DoEvents
    If picDate.Visible = False Then Selection_Change
End Sub
Private Sub mthView_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    lblDate.Caption = Format(mthView.SelStart, "DD MMM YYYY") & " to " & Format(mthView.SelEnd, "DD MMM YYYY")
End Sub

