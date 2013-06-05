VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTransfers 
   Caption         =   "Stock Transfers"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   13725
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   8100
      Index           =   1
      Left            =   6960
      TabIndex        =   1
      Top             =   690
      Width           =   6675
      _cx             =   11774
      _cy             =   14287
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorAlternate=   16381166
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox picPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   420
         ScaleHeight     =   2805
         ScaleWidth      =   5835
         TabIndex        =   6
         Top             =   2430
         Width           =   5835
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   330
            Top             =   150
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "Show out of Stock Products in Transfering Location."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   690
            TabIndex        =   8
            Top             =   1470
            Width           =   4755
         End
         Begin VB.CheckBox chkZero 
            Caption         =   "Show out of Stock Products in Receiving Location."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   690
            TabIndex        =   7
            Top             =   1860
            Value           =   1  'Checked
            Width           =   4755
         End
         Begin VB.Label lblPrompt 
            Alignment       =   2  'Center
            Caption         =   "Select a Location to Transfer to. >>>>"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   420
            TabIndex        =   9
            Top             =   660
            Width           =   4965
         End
         Begin MSForms.Image Image7 
            Height          =   2805
            Left            =   0
            Top             =   0
            Width           =   5835
            BorderStyle     =   0
            SpecialEffect   =   6
            Size            =   "10292;4948"
         End
      End
      Begin VB.Label lblTransfer 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   810
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   8070
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   690
      Width           =   6825
      _cx             =   12039
      _cy             =   14235
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   0
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
      Rows            =   1
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSForms.ComboBox cmbTransRec 
      Height          =   375
      Left            =   10320
      TabIndex        =   5
      Top             =   180
      Width           =   3285
      VariousPropertyBits=   746588185
      DisplayStyle    =   7
      Size            =   "5794;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbTransLoc 
      Height          =   375
      Left            =   3450
      TabIndex        =   4
      Top             =   180
      Width           =   3405
      DisplayStyle    =   7
      Size            =   "6006;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image6 
      Height          =   90
      Left            =   7020
      Top             =   480
      Width           =   2115
      BackColor       =   16761024
      Size            =   "3731;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   6
      Left            =   9180
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   5
      Left            =   9510
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   4
      Left            =   9840
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin VB.Label Label1 
      Caption         =   "Receiving Location."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7110
      TabIndex        =   3
      Top             =   150
      Width           =   3135
   End
   Begin MSForms.Image Image5 
      Height          =   90
      Left            =   120
      Top             =   480
      Width           =   2115
      BackColor       =   16761024
      Size            =   "3731;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   3
      Left            =   2280
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   1
      Left            =   2610
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin MSForms.Image Image4 
      Height          =   90
      Index           =   2
      Left            =   2940
      Top             =   480
      Width           =   285
      BackColor       =   16761024
      Size            =   "503;159"
   End
   Begin VB.Label lblUsers 
      Caption         =   "Transfering Location."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   150
      Width           =   3135
   End
   Begin MSForms.Image Image4 
      Height          =   8145
      Index           =   0
      Left            =   6960
      Top             =   690
      Width           =   6705
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "11827;14367"
   End
   Begin MSForms.Image Image3 
      Height          =   555
      Left            =   6960
      Top             =   90
      Width           =   6705
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "11827;979"
   End
   Begin MSForms.Image Image2 
      Height          =   8145
      Left            =   60
      Top             =   690
      Width           =   6855
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12091;14367"
   End
   Begin MSForms.Image Image1 
      Height          =   555
      Left            =   60
      Top             =   90
      Width           =   6855
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12091;979"
   End
End
Attribute VB_Name = "frmTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTransLoc_Change()
    On Error Resume Next
    grdMain(0).Rows = 1
    If cmbTransLoc.Text <> "<Select a Location>" Then
        cmbTransRec.Enabled = True
        If Trim(cmbTransLoc.Text) <> "" Then
            ActiveReadServer "Select Stock_on_Hand,Product_Code,(Select Ave_Cost from Products where  Products.Product_code=Quantities.Product_Code) as Ave_Cost,(Select Department_No from Products where  Products.Product_code=Quantities.Product_Code) as Department_No,(Select  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
            "+ Unit_of_Measure END from Products where Products.Product_code=Quantities.Product_Code) as Description from Quantities where Location_No = " & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & " order by Description"
            While Not rs.EOF
                grdMain(0).Rows = grdMain(0).Rows + 1
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 0) = rs.Fields("Product_code")
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 1) = rs.Fields("Description")
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 2) = Round(Val(rs.Fields("Stock_on_Hand") & ""), 3)
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 3) = "0"
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 4) = rs.Fields("Department_No")
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 5) = rs.Fields("Ave_Cost")
                rs.MoveNext
            Wend
            rs.Close
            If grdMain(0).Rows > 1 Then
                grdMain(0).Row = 1
            End If
            cmbTransRec.Tag = "1"
            cmbTransRec.Clear
            If cmbTransLoc.Text <> "<Select a Location>" Then
                ActiveReadServer "Select Location_No,Loc_Name from Locations where Location_no <> " & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & " order by Location_no"
                cmbTransRec.AddItem "<Select a Location>"
                While Not rs.EOF
                    cmbTransRec.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
                    rs.MoveNext
                Wend
                rs.Close
                cmbTransRec.Text = "<Select a Location>"
            End If
            cmbTransRec.Tag = ""
            Timer1.Enabled = True
            picPrompt.Visible = True
        End If
    Else
        cmbTransRec.Enabled = False
    End If
    On Error GoTo 0
End Sub
Private Sub cmbTransRec_Change()
    grdMain(1).Rows = 1
    picPrompt.Visible = False
    Timer1.Enabled = False
    grdMain(0).Enabled = True
    If cmbTransRec.Tag = "" Then
        If chkZero(0) = 0 Then
            grdMain(0).Rows = 1
            ActiveReadServer "Select (Select Ave_Cost from Products where  Products.Product_code=Quantities.Product_Code) as Ave_Cost,(Select Department_No from Products where  Products.Product_code=Quantities.Product_Code) as Department_No,Stock_on_Hand,Product_Code,(Select  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
            "+ Unit_of_Measure END from Products where Products.Product_code=Quantities.Product_Code) as Description from Quantities where Stock_on_Hand<>0 and Location_No = " & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & " order by Description"
            While Not rs.EOF
                grdMain(0).Rows = grdMain(0).Rows + 1
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 0) = rs.Fields("Product_code")
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 1) = rs.Fields("Description") & ""
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 2) = Round(Val(rs.Fields("Stock_on_Hand") & ""), 3)
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 3) = "0"
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 4) = rs.Fields("Department_No") & ""
                grdMain(0).TextMatrix(grdMain(0).Rows - 1, 5) = Val(rs.Fields("Ave_Cost") & "")
                rs.MoveNext
            Wend
            rs.Close
            If grdMain(0).Rows > 1 Then
                grdMain(0).Row = 1
            End If
        End If
        If chkZero(1) = 1 Then
            grdMain(1).Rows = 1
            ActiveReadServer "Select Stock_on_Hand,Product_Code,(Select  CASE Unit_Size WHEN 0 THEN Products.Description + ' ' + Unit_of_Measure ELSE Products.Description + ' ' + CONVERT(nvarchar(20), Unit_Size) " & _
            "+ Unit_of_Measure END from Products where Products.Product_code=Quantities.Product_Code) as Description from Quantities where (Stock_on_Hand = 0 or Stock_on_Hand < 0) and Location_No = " & Val(Mid(cmbTransRec.Text, 1, InStr(cmbTransRec.Text, "-") - 1)) & _
            "and Product_Code in (Select Product_Code from Quantities where Location_No =" & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & ") order by Description"
            While Not rs.EOF
                grdMain(1).Rows = grdMain(1).Rows + 1
                grdMain(1).TextMatrix(grdMain(1).Rows - 1, 0) = rs.Fields("Product_code")
                grdMain(1).TextMatrix(grdMain(1).Rows - 1, 1) = rs.Fields("Description") & ""
                grdMain(1).TextMatrix(grdMain(1).Rows - 1, 2) = Round(Val(rs.Fields("Stock_on_Hand") & ""), 3)
                grdMain(1).TextMatrix(grdMain(1).Rows - 1, 3) = "0"
                rs.MoveNext
            Wend
            rs.Close
        End If
        cmbTransRec.Enabled = False
        grdMain(0).SetFocus
        grdMain(1).Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    frmMain.Toolbar1.Buttons(11).Enabled = False
    frmMain.Toolbar1.Buttons(12).Enabled = False
End Sub
Private Sub Form_Load()
    lblTransfer.Caption = ""
    frmMain.Toolbar1.Buttons(2).Caption = "Accept"
    frmMain.Toolbar1.Buttons(2).Enabled = False
    frmMain.Toolbar1.Buttons(3).Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    cmbTransLoc.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    cmbTransLoc.AddItem "<Select a Location>"
    While Not rs.EOF
        cmbTransLoc.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmbTransLoc.Text = "<Select a Location>"
    cmbTransRec.Tag = "1"
    cmbTransRec.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    cmbTransRec.AddItem "<Select a Location>"
    While Not rs.EOF
        cmbTransRec.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmbTransRec.Text = "<Select a Location>"
    cmbTransRec.Tag = ""
    For i = 0 To 1
        grdMain(i).Rows = 1
        grdMain(i).Cols = 6
        grdMain(i).TextMatrix(0, 0) = "Product Code"
        grdMain(i).TextMatrix(0, 1) = "Description"
        grdMain(i).TextMatrix(0, 2) = "Stock on Hand"
        grdMain(i).TextMatrix(0, 3) = "   Transfer Qty"
        grdMain(i).RowHeight(0) = 550
        grdMain(i).ColWidth(0) = grdMain(i).Width * 0.2
        grdMain(i).ColWidth(1) = grdMain(i).Width * 0.5
        grdMain(i).ColWidth(2) = grdMain(i).Width * 0.15
        grdMain(i).ColWidth(3) = grdMain(i).Width * 0.15
        grdMain(i).ColAlignment(0) = flexAlignLeftCenter
        grdMain(i).ColAlignment(1) = flexAlignLeftCenter
        grdMain(i).ColAlignment(2) = flexAlignRightCenter
        grdMain(i).ColAlignment(3) = flexAlignRightCenter
        grdMain(i).ColHidden(4) = True
        grdMain(i).ColHidden(5) = True
    Next i
End Sub
Public Sub Accept_Transfer()
    frmMain.Toolbar1.Tag = ""
    Select Case MsgBox("Do you want to Accept the Current Transfer?", vbYesNoCancel, "HeroPOS")
        Case vbYes
            Screen.MousePointer = 11
            If Val(lblTransfer.Caption) <> 0 Then
                ActiveUpdateServer "Delete from Transfer_Listing where Transfer_No = " & Val(lblTransfer.Caption)
            Else
                ActiveReadServer1 "Select isnull(max(Transfer_No),0) + 1 as Transfer_No from Transfer_Journal"
                lblTransfer.Caption = Format(rs1.Fields("Transfer_No"), "000000")
                rs1.Close
            End If
            DoEvents
            NewQuantity = 0
            For i = 1 To grdMain(1).Rows - 1
                If grdMain(1).ValueMatrix(i, 3) <> 0 Then
                    ActiveUpdateServer "INSERT INTO Transfer_Journal (Workstation_No, User_No, Trans_Location_No,Rec_Location_No, Transfer_No, Product_Code, Department_No, Qty,Date_Time,Ave_Cost)" & _
                    " VALUES (" & Workstation_No & "," & UserRecord.User_Number & "," & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & "," & Val(Mid(cmbTransRec.Text, 1, InStr(cmbTransRec.Text, "-") - 1)) & "," & Val(lblTransfer.Caption) & ",'" & grdMain(1).TextMatrix(i, 0) & "','" & grdMain(1).TextMatrix(i, 4) & "'," & grdMain(1).TextMatrix(i, 3) & ",Getdate()," & grdMain(1).TextMatrix(i, 5) & ")"
                    
                    ActiveReadServer "Select Stock_on_Hand from Quantities where Product_Code = '" & grdMain(1).TextMatrix(i, 0) & "' and Location_No = " & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1))
                    If rs.RecordCount > 0 Then
                        ActiveUpdateServer "Update Quantities Set Stock_on_Hand = Stock_on_Hand - " & grdMain(1).TextMatrix(i, 3) & " where Product_Code = '" & grdMain(1).TextMatrix(i, 0) & "' and Location_No = " & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1))
                    Else
                        ActiveUpdateServer "INSERT INTO Quantities (Product_Code,Location_No,Stock_on_Hand) values ('" & grdMain(1).TextMatrix(i, 0) & "'," & Val(Mid(cmbTransLoc.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & "," & grdMain(1).ValueMatrix(i, 0) * -1 & ")"
                    End If
                    rs.Close
                    Loc_Type = 0
                    ActiveReadServer1 "Select Loc_Type from Locations where Location_no = " & Val(Mid(cmbTransRec.Text, 1, InStr(cmbTransRec.Text, "-") - 1))
                    If rs1.RecordCount > 0 Then
                        Loc_Type = rs1.Fields("Loc_Type")
                    End If
                    rs1.Close
                    If Loc_Type <> 3 Then
                        ActiveReadServer "Select Stock_on_Hand from Quantities where Product_Code = '" & grdMain(1).TextMatrix(i, 0) & "' and Location_No = " & Val(Mid(cmbTransRec.Text, 1, InStr(cmbTransRec.Text, "-") - 1))
                        If rs.RecordCount > 0 Then
                            ActiveUpdateServer "Update Quantities Set Stock_on_Hand = Stock_on_Hand + " & grdMain(1).TextMatrix(i, 3) & " where Product_Code = '" & grdMain(1).TextMatrix(i, 0) & "' and Location_No = " & Val(Mid(cmbTransRec.Text, 1, InStr(cmbTransRec.Text, "-") - 1))
                        Else
                            ActiveUpdateServer "INSERT INTO Quantities (Product_Code,Location_No,Stock_on_Hand) values ('" & grdMain(1).TextMatrix(i, 0) & "'," & Val(Mid(cmbTransRec.Text, 1, InStr(cmbTransRec.Text, "-") - 1)) & "," & grdMain(1).ValueMatrix(i, 3) & ")"
                        End If
                        rs.Close
                    End If
                End If
            Next i
            MsgBox "Transfer no: " & lblTransfer.Caption & " Accepted Successfully", vbInformation, "HeroPOS"
            Screen.MousePointer = 0
            frmMain.Toolbar1.Tag = lblTransfer.Caption
            Form_Load
        Case vbNo
            frmMain.picProdBar.Visible = False
            Unload frmTransfer
            DoEvents
            frmMain.cmdBar(7).Enabled = True
            frmDetails.Show
        Case Else
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Toolbar1.Buttons(11).Enabled = True
    frmMain.Toolbar1.Buttons(12).Enabled = True
End Sub

Private Sub grdMain_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If grdMain(0).ValueMatrix(Row, 2) <> 0 Then
    frmMain.Toolbar1.Buttons(2).Enabled = True
     End If
    Select Case Index
    
   
        Case 0
'            If grdMain(0).ValueMatrix(Row, 2) < 0 Then
'                grdMain(0).TextMatrix(Row, 3) = "0"
'                Exit Sub
'            End If
'            If grdMain(0).ValueMatrix(Row, 2) - grdMain(0).ValueMatrix(Row, 3) < 0 Then
'                grdMain(0).TextMatrix(Row, 3) = grdMain(0).ValueMatrix(Row, 2)
'            End If
            Select Case grdMain(1).FindRow(grdMain(0).TextMatrix(Row, 0), 0, 0)
                Case -1
                    grdMain(1).Rows = grdMain(1).Rows + 1
                    grdMain(1).Row = grdMain(1).Rows - 1
                    grdMain(1).TextMatrix(grdMain(1).Row, 0) = grdMain(0).TextMatrix(Row, 0)
                    grdMain(1).TextMatrix(grdMain(1).Row, 1) = grdMain(0).TextMatrix(Row, 1)
                    ActiveReadServer "Select Stock_on_Hand from Quantities where Location_No = " & Val(Mid(cmbTransRec.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & " and Product_Code = '" & grdMain(0).TextMatrix(Row, 0) & "'"
                    If rs.RecordCount > 0 Then
                        grdMain(1).TextMatrix(grdMain(1).Row, 2) = rs.Fields("Stock_on_Hand")
                    Else
                        grdMain(1).TextMatrix(grdMain(1).Row, 2) = 0
                    End If
                    rs.Close
                    grdMain(1).TextMatrix(grdMain(1).Row, 3) = grdMain(0).TextMatrix(Row, 3)
                    grdMain(1).TextMatrix(grdMain(1).Row, 4) = grdMain(0).TextMatrix(Row, 4)
                    grdMain(1).TextMatrix(grdMain(1).Row, 5) = grdMain(0).TextMatrix(Row, 5)
                    grdMain(0).Cell(flexcpFontBold, Row, 0, Row, 2) = True
                    grdMain(0).Cell(flexcpForeColor, Row, 0, Row, 2) = &HFF0000
                    grdMain(0).Cell(flexcpFontBold, Row, 3, Row, 3) = True
                    grdMain(0).Cell(flexcpForeColor, Row, 3, Row, 3) = &HC0&
                    grdMain(1).ShowCell grdMain(1).Row, 0
                Case Else
                    grdMain(1).Row = grdMain(1).FindRow(grdMain(0).TextMatrix(Row, 0), 0, 0)
                    grdMain(1).TextMatrix(grdMain(1).Row, 0) = grdMain(0).TextMatrix(Row, 0)
                    grdMain(1).TextMatrix(grdMain(1).Row, 1) = grdMain(0).TextMatrix(Row, 1)
                    ActiveReadServer "Select Stock_on_Hand from Quantities where Location_No = " & Val(Mid(cmbTransRec.Text, 1, InStr(cmbTransLoc.Text, "-") - 1)) & " and Product_Code = '" & grdMain(0).TextMatrix(Row, 0) & "'"
                    If rs.RecordCount > 0 Then
                        grdMain(1).TextMatrix(grdMain(1).Row, 2) = rs.Fields("Stock_on_Hand")
                    Else
                        grdMain(1).TextMatrix(grdMain(1).Row, 2) = 0
                    End If
                    rs.Close
                    grdMain(1).TextMatrix(grdMain(1).Row, 3) = grdMain(0).TextMatrix(Row, 3)
                    grdMain(1).TextMatrix(grdMain(1).Row, 4) = grdMain(0).TextMatrix(Row, 4)
                    grdMain(1).TextMatrix(grdMain(1).Row, 5) = grdMain(0).TextMatrix(Row, 5)
                    grdMain(0).Cell(flexcpFontBold, Row, 0, Row, 2) = True
                    grdMain(0).Cell(flexcpForeColor, Row, 0, Row, 2) = &HFF0000
                    grdMain(0).Cell(flexcpFontBold, Row, 3, Row, 3) = True
                    grdMain(0).Cell(flexcpForeColor, Row, 3, Row, 3) = &HC0&
                    grdMain(1).ShowCell grdMain(1).Row, 0
            End Select
    End Select
    noQty = True
    For i = 1 To grdMain(1).Rows - 1
        If grdMain(1).ValueMatrix(i, 3) <> 0 Then
            frmMain.Toolbar1.Buttons(2).Enabled = True
            noQty = False
            Exit For
        End If
    Next i
    frmMain.picBar.Tag = "Transfer"
    If noQty = True Then frmMain.Toolbar1.Buttons(2).Enabled = False
End Sub
Private Sub grdMain_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If Index = 0 Then
        If grdMain(1).FindRow(grdMain(0).TextMatrix(NewRow, 0), 0, 0) <> -1 Then
            grdMain(1).Row = grdMain(1).FindRow(grdMain(0).TextMatrix(NewRow, 0), 0, 0)
        Else
            grdMain(1).Row = 0
        End If
        
    End If
    ActiveReadServer "Select Landed_Cost,Selling_Price,Sales_Tax,(Select Price2 from Product_Prices where " & _
    "Products.Product_code = Product_Prices.Product_Code) as Retail_Price From Products where Product_code = '" & grdMain(0).TextMatrix(NewRow, 0) & "'"
    If rs.RecordCount > 0 Then
        If Branch_Type = 10 Then
            frmMain.stbBar.Panels(3).Text = "Cost Price: " & Format(Val(rs.Fields("Landed_Cost") & ""), "0.00") & _
            "   Selling Price: " & Format(Val(rs.Fields("Selling_Price") & ""), "0.00") & _
            "   Retail Price: " & Format(Val(rs.Fields("Retail_Price") & ""), "0.00")
            DoEvents
        Else
            frmMain.stbBar.Panels(3).Text = "Cost Price: " & Format(Val(rs.Fields("Landed_Cost") & ""), "0.00") & _
            "   Selling Price: " & Format(Val(rs.Fields("Selling_Price") & ""), "0.00")
            DoEvents
        End If
    End If
    rs.Close
    On Error GoTo 0
End Sub

Private Sub grdMain_EnterCell(Index As Integer)
    Select Case Index
        Case 0
            Select Case grdMain(0).Col
                Case 0
                    grdMain(0).Editable = flexEDNone
                Case 1
                    grdMain(0).Editable = flexEDNone
                Case 2
                    grdMain(0).Editable = flexEDNone
                    grdMain(0).Col = 3
                Case 3
                    If grdMain(0).ValueMatrix(grdMain(0).Row, 2) = 0 Then
                        grdMain(0).Editable = flexEDNone
                    Else
                        If picPrompt.Visible = True Then
                            grdMain(0).Editable = flexEDNone
                        Else
                            grdMain(0).Editable = flexEDKbdMouse
                        End If
                    End If
            End Select
        Case 1
            grdMain(1).Editable = flexEDNone
    End Select
End Sub
Private Sub grdMain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And grdMain(0).Col = 3 Then
        Select Case KeyCode
            Case 13 'Enter
            Case 45, 48 To 57, 96 To 105, 109, 110, 189
                If grdMain(0).ValueMatrix(grdMain(0).Row, 2) <> 0 And picPrompt.Visible = False Then grdMain(0).EditCell
            Case 38 'up
            Case 40 'down
            Case 37
                grdMain(0).Col = 1
        End Select
    Else
        If grdMain(1).ValueMatrix(grdMain(1).Row, 3) <> 0 And KeyCode = 46 Then
            grdMain(0).TextMatrix(grdMain(0).FindRow(grdMain(1).TextMatrix(grdMain(1).Row, 0), 0, 0), 3) = "0"
            grdMain(0).Cell(flexcpFontBold, grdMain(0).FindRow(grdMain(1).TextMatrix(grdMain(1).Row, 0), 0, 0), 0, grdMain(0).FindRow(grdMain(1).TextMatrix(grdMain(1).Row, 0), 0, 0), 3) = False
            grdMain(0).Cell(flexcpForeColor, grdMain(0).FindRow(grdMain(1).TextMatrix(grdMain(1).Row, 0), 0, 0), 0, grdMain(0).FindRow(grdMain(1).TextMatrix(grdMain(1).Row, 0), 0, 0), 3) = grdMain(0).ForeColor
            grdMain(1).RemoveItem (grdMain(1).Row)
            noQty = True
            For i = 1 To grdMain(1).Rows - 1
                If grdMain(1).ValueMatrix(i, 3) <> 0 Then
                    frmMain.Toolbar1.Buttons(2).Enabled = True
                    noQty = False
                    Exit For
                End If
            Next i
            If noQty = True Then frmMain.Toolbar1.Buttons(2).Enabled = False
        End If
    End If
End Sub
Private Sub grdMain_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 And grdMain(0).Col = 3 Then
        If InStr(grdMain(0).EditText, ".") <> 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If
        Select Case KeyAscii
            Case 8, 13, 27, 45, 46, 48 To 57
            Case Else
                If Col = 0 Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub Timer1_Timer()
    Select Case lblPrompt.ForeColor
        Case &H80000012
            lblPrompt.ForeColor = &HC00000
        Case &HC00000
            lblPrompt.ForeColor = &H80000012
    End Select
End Sub
Public Sub DeleteTransfer()
    Select Case MsgBox("Are you sure you want to Delete the Current Transfer", vbYesNo, "HeroPOS")
        Case vbYes
            ActiveUpdateServer "Delete from Transfer_Listing where Transfer_No = " & Val(lblTransfer.Caption)
            ActiveUpdateServer "Delete from Transfer_Journal where Transfer_No = " & Val(lblTransfer.Caption)
            frmMain.picProdBar.Visible = False
            Unload frmGRV
            DoEvents
            frmMain.cmdBar(7).Enabled = True
            frmDetails.Show
        Case Else
    End Select
End Sub

