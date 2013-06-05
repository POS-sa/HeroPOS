VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmGP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Stock Consumption Analysis..."
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4380
   FillColor       =   &H00F8D9CF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid grdStock 
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4260
      _cx             =   7514
      _cy             =   2037
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   15329975
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   8421504
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   380
      RowHeightMax    =   0
      ColWidthMin     =   2500
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGP.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin btButtonEx.ButtonEx cmdSupplier 
      Height          =   345
      Left            =   3075
      TabIndex        =   1
      ToolTipText     =   " Click to Search.... "
      Top             =   5160
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
      Style           =   1
      ShowFocus       =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid grdRev 
      Height          =   1350
      Left            =   60
      TabIndex        =   2
      Top             =   1770
      Width           =   4245
      _cx             =   7488
      _cy             =   2381
      Appearance      =   2
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
      BackColorSel    =   12582912
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   1800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGP.frx":0054
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   End
   Begin VSFlex8Ctl.VSFlexGrid grdRev1 
      Height          =   1350
      Left            =   60
      TabIndex        =   3
      Top             =   3720
      Width           =   4245
      _cx             =   7488
      _cy             =   2381
      Appearance      =   2
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
      BackColorSel    =   12582912
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   1800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGP.frx":00CC
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   End
   Begin MSForms.Label Label5 
      Height          =   345
      Left            =   90
      TabIndex        =   5
      Top             =   1320
      Width           =   4185
      ForeColor       =   8421504
      Caption         =   "Theoretical GP"
      Size            =   "7382;609"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   3240
      Width           =   4185
      ForeColor       =   8421504
      Caption         =   "Actual GP"
      Size            =   "7382;661"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image Image7 
      Height          =   495
      Left            =   60
      Top             =   3180
      Width           =   4245
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7488;873"
   End
   Begin MSForms.Image Image2 
      Height          =   435
      Left            =   60
      Top             =   1290
      Width           =   4245
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "7488;767"
   End
End
Attribute VB_Name = "frmGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSupplier_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    grdRev.TextMatrix(0, 0) = "Revenue"
    grdRev.TextMatrix(1, 0) = "-Cost of Sales"
    grdRev.TextMatrix(2, 0) = "=Gross Profit"
    grdRev.TextMatrix(3, 0) = "=Gross Profit%"
    grdRev.TextMatrix(0, 1) = "0.00"
    grdRev.TextMatrix(1, 1) = "0.00"
    grdRev.TextMatrix(2, 1) = "0.00"
    grdRev.TextMatrix(3, 1) = "0%"
    grdRev.Cell(flexcpFontBold, 2, 0, 2, 1) = True
    grdRev.Cell(flexcpFontBold, 3, 0, 3, 1) = True
    grdRev1.TextMatrix(0, 0) = "Revenue"
    grdRev1.TextMatrix(1, 0) = "-Cost of Sales"
    grdRev1.TextMatrix(2, 0) = "=Gross Profit"
    grdRev1.TextMatrix(3, 0) = "=Gross Profit%"
    grdRev1.TextMatrix(0, 1) = "0.00"
    grdRev1.TextMatrix(1, 1) = "0.00"
    grdRev1.TextMatrix(2, 1) = "0.00"
    grdRev1.TextMatrix(3, 1) = "0%"
    grdRev1.Cell(flexcpFontBold, 2, 0, 2, 1) = True
    grdRev1.Cell(flexcpFontBold, 3, 0, 3, 1) = True
    grdRev.Cell(flexcpBackColor, 2, 1, 3, 1) = &HFDE2D9
    grdRev1.Cell(flexcpBackColor, 2, 1, 3, 1) = &HFDE2D9
    grdStock.TextMatrix(0, 0) = "Sales Consumption"
    grdStock.TextMatrix(1, 0) = "Stock Variance"
    grdStock.TextMatrix(2, 0) = "Total Consumption"
    grdStock.TextMatrix(0, 1) = "0.00"
    grdStock.TextMatrix(1, 1) = "0.00"
    grdStock.TextMatrix(2, 1) = "0.00"
    grdStock.Cell(flexcpBackColor, 2, 0, 2, 1) = &HF8D9CF
    If Right(Str(Time_Stop), 2) = "AM" Then
        Selender = DateAdd("d", 1, frmReports.mthViewEnd.Value)
    Else
        Selender = frmReports.mthViewEnd.Value
    End If
    If frmReports.cmb1.Text <> "<All Locations>" Then
        LocString = Mid(frmReports.cmb1.Text, 1, InStr(frmReports.cmb1.Text, "-") - 2)
    Else
        LocString = "%"
    End If
    If frmReports.cmb3.Text = "<All Departments>" Then
        DeptString = "%"
    Else
        If InStr(Mid(frmReports.cmb3.Text, 1, InStrRev(frmReports.cmb3.Text, "-") - 2), "-") = 0 Then
            DeptString = Mid(frmReports.cmb3.Text, 1, InStrRev(frmReports.cmb3.Text, "-") - 2) & "%"
        Else
            DeptString = Mid(frmReports.cmb3.Text, 1, InStrRev(frmReports.cmb3.Text, "-") - 2)
        End If
    End If
    Variance = 0
    ActiveReadServer "Select sum(Variance*Ave_Cost) as Variance" & _
    " from Stock_Take_Variance where " & _
    " Department_No like '" & DeptString & "'" & _
    " and Location_No like '" & LocString & "'" & _
    " and (Date_Time > '" & frmReports.mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    
    If rs.RecordCount > 0 Then
        Variance = Val(rs.Fields("Variance") & "") * -1
    End If
    rs.Close
    grdStock.TextMatrix(1, 1) = Format(Variance, "0.00")
    DoEvents
    
  

    ActiveReadServer "SELECT SUM(Ave_Cost*Qty) AS Ave_Cost" & _
    " From Sales_Journal" & _
    " Where Function_Key=7 and Department_No like '" & DeptString & "'" & _
    " and Location like '" & LocString & "'" & _
    " and (Date_Time > '" & frmReports.mthViewStart.Value & " " & Format(Time_Start, "hh:mm:ss AM/PM") & "' and Date_Time<'" & Selender & " " & Format(Time_Stop, "hh:mm:ss AM/PM") & "')"
    On Error Resume Next
    If rs.RecordCount > 0 Then
        Ave_Cost = Val(rs.Fields("Ave_Cost") & "")
    End If
    rs.Close
    grdStock.TextMatrix(0, 1) = Format(Ave_Cost, "0.00")
    grdStock.TextMatrix(2, 1) = Format(grdStock.ValueMatrix(0, 1) + grdStock.ValueMatrix(1, 1), "0.00")
    DoEvents
    grdRev.TextMatrix(0, 1) = frmReports.grdRev.TextMatrix(6, 1)
    grdRev.TextMatrix(1, 1) = grdStock.TextMatrix(0, 1)
    grdRev1.TextMatrix(1, 1) = grdStock.TextMatrix(2, 1)
    
    grdRev1.TextMatrix(0, 1) = frmReports.grdRev.TextMatrix(6, 1)
    grdRev.TextMatrix(2, 1) = Format((grdRev.ValueMatrix(0, 1) - frmReports.grdTax.ValueMatrix(4, 1)), "0.00") - grdRev.ValueMatrix(1, 1)
    If grdRev1.ValueMatrix(1, 1) > 0 Then grdRev1.TextMatrix(2, 1) = Format((grdRev1.ValueMatrix(0, 1) - frmReports.grdTax.ValueMatrix(4, 1)), "0.00") - grdRev1.ValueMatrix(1, 1)
    If grdRev1.ValueMatrix(1, 1) < 0 Then grdRev1.TextMatrix(2, 1) = Format((grdRev1.ValueMatrix(0, 1) - frmReports.grdTax.ValueMatrix(4, 1)), "0.00") + grdRev1.ValueMatrix(1, 1)
    
    grdRev.TextMatrix(3, 1) = Round((grdRev.TextMatrix(2, 1) / (grdRev.ValueMatrix(0, 1) - frmReports.grdTax.ValueMatrix(4, 1))) * 100, 3) & "%"
    grdRev1.TextMatrix(3, 1) = Round((grdRev1.TextMatrix(2, 1) / (grdRev1.ValueMatrix(0, 1) - frmReports.grdTax.ValueMatrix(4, 1))) * 100, 3) & "%"
End Sub



