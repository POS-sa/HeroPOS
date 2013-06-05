VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStateRun 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Print Statement Run..."
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin btButtonEx.ButtonEx cmdEmptyrprtfldr 
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      ToolTipText     =   " Click to Search.... "
      Top             =   7770
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Empty Rprt Folder"
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
   Begin VB.CheckBox chkAll 
      Caption         =   "Select all"
      Height          =   405
      Left            =   2310
      TabIndex        =   7
      Top             =   7980
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   120
      TabIndex        =   3
      Top             =   7770
      Visible         =   0   'False
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      ToolTipText     =   " Click to Search.... "
      Top             =   8250
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
      Left            =   7080
      TabIndex        =   1
      ToolTipText     =   " Click to Search.... "
      Top             =   8250
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Start"
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
      Height          =   7650
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   9540
      _cx             =   16828
      _cy             =   13494
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
      FormatString    =   $"frmStateRun.frx":0000
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
   Begin MSForms.OptionButton OptionButton2 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   8370
      Width           =   2175
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "3836;503"
      Value           =   "0"
      Caption         =   " Hide 0.00 Balances"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton OptionButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   7980
      Width           =   2175
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "3836;661"
      Value           =   "0"
      Caption         =   "Show 0.00 Balances"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkpdf 
      Height          =   345
      Left            =   2310
      TabIndex        =   4
      Top             =   8370
      Width           =   1335
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2355;609"
      Value           =   "0"
      Caption         =   "Export to PDF"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmStateRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub chkAll_Click()
If chkAll.Value = 1 Then
grdList.Cols = 4
    grdList.TextMatrix(0, 0) = " Debtor No"
    grdList.TextMatrix(0, 1) = " Debtor Name"
    grdList.TextMatrix(0, 2) = " Balance"
    grdList.TextMatrix(0, 3) = " Select"
    grdList.ColAlignment(0) = flexAlignLeftCenter
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.ColDataType(3) = flexDTBoolean
    grdList.ColWidth(0) = grdList.Width * 0.15
    grdList.ColWidth(1) = grdList.Width * 0.51
    grdList.ColWidth(2) = grdList.Width * 0.2
    grdList.Rows = 1
    ActiveReadServer "Select * from Debtors order by Debtor_Name"
    While Not rs.EOF
        grdList.Rows = grdList.Rows + 1
        grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("Debtor_Name")
        grdList.TextMatrix(grdList.Rows - 1, 2) = Format(rs.Fields("Balance"), "0.00")
        grdList.TextMatrix(grdList.Rows - 1, 3) = True
        rs.MoveNext
    Wend
    rs.Close
    Exit Sub
    End If
    grdList.Cols = 4
    grdList.TextMatrix(0, 0) = " Debtor No"
    grdList.TextMatrix(0, 1) = " Debtor Name"
    grdList.TextMatrix(0, 2) = " Balance"
    grdList.TextMatrix(0, 3) = " Select"
    grdList.ColAlignment(0) = flexAlignLeftCenter
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.ColDataType(3) = flexDTBoolean
    grdList.ColWidth(0) = grdList.Width * 0.15
    grdList.ColWidth(1) = grdList.Width * 0.51
    grdList.ColWidth(2) = grdList.Width * 0.2
    grdList.Rows = 1
    ActiveReadServer "Select * from Debtors order by Debtor_Name"
    While Not rs.EOF
        grdList.Rows = grdList.Rows + 1
        grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("Debtor_Name")
        grdList.TextMatrix(grdList.Rows - 1, 2) = Format(rs.Fields("Balance"), "0.00")
        grdList.TextMatrix(grdList.Rows - 1, 3) = ""
        rs.MoveNext
    Wend
    rs.Close
End Sub





Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEmptyrprtfldr_Click()
On Error GoTo Errors
Kill App.Path & "\PDFReports\" & "*.pdf"
Errors:
End Sub

Private Sub cmdOpen_Click()
    If chkpdf.Value = True Then
    If Not regDoes_Key_Exist(HKEY_CURRENT_USER, "Software\CUSTPDF Writer") Then
    Load frmError
    frmError.lblCap.Caption = "Please install or reinstall " & "PDFill PDF Writer" & " first before trying to export to PDF."
    frmError.Show vbModal
    Exit Sub
    End If
    Dim fso As New FileSystemObject
    If fso.FolderExists(App.Path & "\PDFReports\") = False Then fso.CreateFolder (App.Path & "\PDFReports")
    
    PrinttoPDF
    Exit Sub
    End If
    Screen.MousePointer = 11
    ProgressBar1.Max = grdList.Rows - 1
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    For i = 1 To grdList.Rows - 1
        grdList.Row = i
        ProgressBar1.Value = ProgressBar1.Value + 1
        If grdList.ValueMatrix(i, 3) <> 0 Then
            DoEvents
            rptStatementrun.PrintReport False
            DoEvents
            Unload rptStatementrun
        End If
        DoEvents
    Next i
    Screen.MousePointer = 0
    
   MsgBox "Print Run Completed successfully spooled to printer", vbInformation, "HeroPOS"
    With frmGuests
    .timerwindow.Enabled = True
    End With
    DoEvents


  Unload Me
    
End Sub

Private Sub Form_Load()
    grdList.Cols = 4
    grdList.TextMatrix(0, 0) = " Debtor No"
    grdList.TextMatrix(0, 1) = " Debtor Name"
    grdList.TextMatrix(0, 2) = " Balance"
    grdList.TextMatrix(0, 3) = " Select"
    grdList.ColAlignment(0) = flexAlignLeftCenter
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.ColDataType(3) = flexDTBoolean
    grdList.ColWidth(0) = grdList.Width * 0.15
    grdList.ColWidth(1) = grdList.Width * 0.51
    grdList.ColWidth(2) = grdList.Width * 0.2
    grdList.Rows = 1
    ActiveReadServer "Select * from Debtors order by Debtor_Name"
    While Not rs.EOF
        grdList.Rows = grdList.Rows + 1
        grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("Debtor_Name")
        grdList.TextMatrix(grdList.Rows - 1, 2) = Format(rs.Fields("Balance"), "0.00")
        grdList.TextMatrix(grdList.Rows - 1, 3) = ""
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub grdList_EnterCell()
    If grdList.Col = 3 Then
   
        grdList.Editable = flexEDKbdMouse
    Else
        grdList.Editable = flexEDNone
    End If
End Sub


Private Sub OptionButton1_Click()

If chkAll.Value = 0 Then
grdList.Clear
grdList.Cols = 4
    grdList.TextMatrix(0, 0) = " Debtor No"
    grdList.TextMatrix(0, 1) = " Debtor Name"
    grdList.TextMatrix(0, 2) = " Balance"
    grdList.TextMatrix(0, 3) = " Select"
    grdList.ColAlignment(0) = flexAlignLeftCenter
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.ColDataType(3) = flexDTBoolean
    grdList.ColWidth(0) = grdList.Width * 0.15
    grdList.ColWidth(1) = grdList.Width * 0.51
    grdList.ColWidth(2) = grdList.Width * 0.2
    grdList.Rows = 1
    ActiveReadServer "Select * from Debtors order by Debtor_Name"
    While Not rs.EOF
        grdList.Rows = grdList.Rows + 1
        grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("Debtor_Name")
        grdList.TextMatrix(grdList.Rows - 1, 2) = Format(rs.Fields("Balance"), "0.00")
        grdList.TextMatrix(grdList.Rows - 1, 3) = ""
        rs.MoveNext
    Wend
    rs.Close
    Exit Sub
    End If
    grdList.Clear
grdList.Cols = 4
    grdList.TextMatrix(0, 0) = " Debtor No"
    grdList.TextMatrix(0, 1) = " Debtor Name"
    grdList.TextMatrix(0, 2) = " Balance"
    grdList.TextMatrix(0, 3) = " Select"
    grdList.ColAlignment(0) = flexAlignLeftCenter
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.ColDataType(3) = flexDTBoolean
    grdList.ColWidth(0) = grdList.Width * 0.15
    grdList.ColWidth(1) = grdList.Width * 0.51
    grdList.ColWidth(2) = grdList.Width * 0.2
    grdList.Rows = 1
    ActiveReadServer "Select * from Debtors order by Debtor_Name"
    While Not rs.EOF
        grdList.Rows = grdList.Rows + 1
        grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("Debtor_Name")
        grdList.TextMatrix(grdList.Rows - 1, 2) = Format(rs.Fields("Balance"), "0.00")
        grdList.TextMatrix(grdList.Rows - 1, 3) = 1
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub OptionButton2_Click()
If chkAll.Value = 0 Then
grdList.Clear
grdList.Cols = 4
    grdList.TextMatrix(0, 0) = " Debtor No"
    grdList.TextMatrix(0, 1) = " Debtor Name"
    grdList.TextMatrix(0, 2) = " Balance"
    grdList.TextMatrix(0, 3) = " Select"
    grdList.ColAlignment(0) = flexAlignLeftCenter
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.ColDataType(3) = flexDTBoolean
    grdList.ColWidth(0) = grdList.Width * 0.15
    grdList.ColWidth(1) = grdList.Width * 0.51
    grdList.ColWidth(2) = grdList.Width * 0.2
    grdList.Rows = 1
    ActiveReadServer "Select * from Debtors Where (Balance > 0) order by Debtor_Name "
    While Not rs.EOF
        grdList.Rows = grdList.Rows + 1
        grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("Debtor_Name")
        grdList.TextMatrix(grdList.Rows - 1, 2) = Format(rs.Fields("Balance"), "0.00")
        grdList.TextMatrix(grdList.Rows - 1, 3) = ""
        rs.MoveNext
    Wend
    rs.Close
    Exit Sub
    End If
    grdList.Clear
grdList.Cols = 4
    grdList.TextMatrix(0, 0) = " Debtor No"
    grdList.TextMatrix(0, 1) = " Debtor Name"
    grdList.TextMatrix(0, 2) = " Balance"
    grdList.TextMatrix(0, 3) = " Select"
    grdList.ColAlignment(0) = flexAlignLeftCenter
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.ColDataType(3) = flexDTBoolean
    grdList.ColWidth(0) = grdList.Width * 0.15
    grdList.ColWidth(1) = grdList.Width * 0.51
    grdList.ColWidth(2) = grdList.Width * 0.2
    grdList.Rows = 1
    ActiveReadServer "Select * from Debtors Where (Balance > 0) order by Debtor_Name "
    While Not rs.EOF
        grdList.Rows = grdList.Rows + 1
        grdList.TextMatrix(grdList.Rows - 1, 0) = rs.Fields("Debtor_No")
        grdList.TextMatrix(grdList.Rows - 1, 1) = rs.Fields("Debtor_Name")
        grdList.TextMatrix(grdList.Rows - 1, 2) = Format(rs.Fields("Balance"), "0.00")
        grdList.TextMatrix(grdList.Rows - 1, 3) = 1
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub PrinttoPDF()
Dim Counter As Integer
Screen.MousePointer = 11
    'ProgressBar1.Max = grdList.Rows - 1
    
    Dim Printdefdevice As String
    Printdefdevice = Printer.DeviceName
    Dim w As New WshNetwork
    Dim Ustring As String
    Dim Usejobnames As String
    Usejobnames = "1"
    w.SetDefaultPrinter (Pdfprinter)
    Set w = Nothing
    regCreate_Key_Value HKEY_CURRENT_USER, "Software\CUSTPDF Writer", "UseJobName", Usejobnames, True
    Call Update(App.Path & Pdffolder)
  
    ProgressBar1.Max = grdList.Rows - 1
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    For i = 1 To grdList.Rows - 1
 
        Ustring = grdList.TextMatrix(i, 0) & grdList.TextMatrix(i, 1)
        PDFAPPNAME = Ustring

        grdList.Row = i
        ProgressBar1.Value = ProgressBar1.Value + 1
        If grdList.ValueMatrix(i, 3) <> 0 Then
        
            DoEvents
            rptStatementrun.PrintReport False
            DoEvents
            Unload rptStatementrun
        End If
        DoEvents
    Next i
    Screen.MousePointer = 0
    MsgBox "Export Run Completed successfully spooled to PDF printer", vbInformation, "HeroPOS"
        With frmGuests
   .timerwindow.Enabled = True
    End With
    DoEvents
    w.SetDefaultPrinter (Printdefdevice)
    
    Unload Me
End Sub
