VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmWage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " User Wage Sheet..."
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   3765
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   3765
   End
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   345
      Left            =   3570
      TabIndex        =   0
      Top             =   4200
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
   Begin btButtonEx.ButtonEx cmdPrint 
      Height          =   345
      Left            =   2310
      TabIndex        =   5
      Top             =   4200
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Print"
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
   Begin VSFlex8Ctl.VSFlexGrid grdRev 
      Height          =   3000
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   4575
      _cx             =   8070
      _cy             =   5292
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      Rows            =   9
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   2300
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmWage.frx":0000
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
   Begin MSForms.Image Image1 
      Height          =   3165
      Left            =   60
      Top             =   990
      Width           =   4725
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "8334;5583"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User: "
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   210
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dated: "
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   570
      Width           =   705
   End
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   780
      Top             =   540
      Width           =   3915
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "6906;609"
   End
   Begin MSForms.Image Image13 
      Height          =   345
      Left            =   780
      Top             =   150
      Width           =   3915
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "6906;609"
   End
   Begin MSForms.Image Image4 
      Height          =   915
      Left            =   60
      Top             =   60
      Width           =   4725
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "8334;1614"
   End
End
Attribute VB_Name = "frmWage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
On Error GoTo trap
With frmSlipDetails
    PrintErr = 0
    Slip_Port = ""
    
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As #filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(69) & Chr(1);
    End If
    For i = 1 To .grdHead.Rows - 1
        If Trim(.grdHead.TextMatrix(i, 0)) <> "" Then
            Select Case Trim(.grdHead.TextMatrix(i, 1))
                Case "Line Feeds"
                    If Slip_Printer_Type = 0 Then
                        Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdHead.TextMatrix(i, 0)));
                    End If
                Case Else
                    Select Case .grdHead.TextMatrix(i, 2)
                        Case "Left"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                        Case "Centre"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                        Case "Right"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                    End Select
                    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                    Print #filenum, Chr(27) & Chr(33) & Chr(0);
                    Select Case Trim(.grdHead.TextMatrix(i, 1))
                        Case ""
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Narrow Font (Dark)"
                            If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font"
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Normal Font (Dark)"
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Double Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(16);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case "Big Font (Dark)"
                            Print #filenum, Chr(27) & Chr(33) & Chr(48);
                            If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                        Case Else
                            Print #filenum, .grdHead.TextMatrix(i, 0)
                    End Select
            End Select
        End If
    Next i
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "USER WAGE SHEET"
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    If Slip_Printer_Type = 0 Then
        Print #filenum, Chr(27) & Chr(77) & Chr(49);
        Print #filenum, String(40, "=")
    Else
        Print #filenum, String(33, "=")
    End If
    If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, UCase(Trim(txtUser.Text))
    Print #filenum, UCase(txtDate.Text)
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(51) & Chr(18);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CASH TOTAL:  " & Chr(179) & String(14 - Len(grdRev.TextMatrix(0, 1)), Chr(32)) & grdRev.TextMatrix(0, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CARD TOTAL:  " & Chr(179) & String(14 - Len(grdRev.TextMatrix(1, 1)), Chr(32)) & grdRev.TextMatrix(1, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "VOUCHER TOTAL:  " & Chr(179) & String(14 - Len(grdRev.TextMatrix(2, 1)), Chr(32)) & grdRev.TextMatrix(2, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CHARGE TOTAL:  " & Chr(179) & String(14 - Len(grdRev.TextMatrix(3, 1)), Chr(32)) & grdRev.TextMatrix(3, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, String(33, "-")
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "TOTAL (INCL):  " & Chr(179) & String(14 - Len(grdRev.TextMatrix(4, 1)), Chr(32)) & grdRev.TextMatrix(4, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "      TAX:  " & Chr(179) & String(14 - Len(grdRev.TextMatrix(5, 1)), Chr(32)) & grdRev.TextMatrix(5, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "TOTAL (EXCL):  " & Chr(179) & String(14 - Len(grdRev.TextMatrix(6, 1)), Chr(32)) & grdRev.TextMatrix(6, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "COMMISSION %:  " & Chr(179) & String(14 - Len(grdRev.TextMatrix(7, 1)), Chr(32)) & grdRev.TextMatrix(7, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "COMMISSION DUE: " & Chr(179) & String(14 - Len(grdRev.TextMatrix(8, 1)), Chr(32)) & grdRev.TextMatrix(8, 1) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, String(33, "-")
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, "PRESENTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "ACCEPTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "DATED:"
    Print #filenum, ""
    Print #filenum, Chr(27) & Chr(50);
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(100) & Chr(7);
    Print #filenum, Chr(27) & Chr(64);
    Print #filenum, Chr(27) & Chr(69) & Chr(1);
    
    
    
    
    
    Print #filenum, Chr(29) & "V" & Chr(49);
    Close #filenum
End With
On Error GoTo 0
Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        Dim x As Printer
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    frmError.Show vbModal
    Close #filenum
    On Error GoTo 0
End Sub
Private Sub Form_Load()
    grdRev.TextMatrix(0, 0) = "Cash Sales:"
    grdRev.TextMatrix(1, 0) = "Card Sales:"
    grdRev.TextMatrix(2, 0) = "Voucher Sales:"
    grdRev.TextMatrix(3, 0) = "Charge Sales:"
    grdRev.TextMatrix(4, 0) = "Total Sales (Incl):"
    grdRev.TextMatrix(5, 0) = "Tax:"
    grdRev.TextMatrix(6, 0) = "Total Sales (Excl)"
    grdRev.TextMatrix(7, 0) = "Commission %:"
    grdRev.TextMatrix(8, 0) = "Commission Due:"
    grdRev.Cell(flexcpFontBold, 4, 0, 4, 1) = True
    grdRev.Cell(flexcpBackColor, 4, 1, 4, 1) = &HE0E0E0
    grdRev.Cell(flexcpFontBold, 6, 0, 6, 1) = True
    grdRev.Cell(flexcpBackColor, 6, 1, 6, 1) = &HE0E0E0
    grdRev.Cell(flexcpFontBold, 8, 0, 8, 1) = True
    grdRev.Cell(flexcpBackColor, 8, 1, 8, 1) = &HE0E0E0
    txtUser.Text = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 0)
    txtDate.Text = frmReports.lblDate
    grdRev.TextMatrix(0, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 1)
    grdRev.TextMatrix(1, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 2)
    grdRev.TextMatrix(2, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 3)
    grdRev.TextMatrix(3, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 4)
    grdRev.TextMatrix(4, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 5)
    grdRev.TextMatrix(5, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 6)
    grdRev.TextMatrix(6, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 7)
    grdRev.TextMatrix(7, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 8)
    grdRev.TextMatrix(8, 1) = frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 9)
End Sub
