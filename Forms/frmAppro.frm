VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAppro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appro's"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   Icon            =   "frmAppro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   220
      Width           =   4365
   End
   Begin VB.TextBox txtTel 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   550
      Width           =   2385
   End
   Begin VB.TextBox txtId 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   880
      Width           =   2385
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   4530
      Left            =   60
      TabIndex        =   0
      Top             =   2610
      Width           =   9225
      _cx             =   16272
      _cy             =   7990
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAppro.frx":000C
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Top             =   5670
         Width           =   1005
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
         TabIndex        =   5
         Top             =   11700
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   345
      Left            =   6810
      TabIndex        =   4
      Top             =   7230
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
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdOk 
      Height          =   345
      Left            =   5550
      TabIndex        =   3
      ToolTipText     =   " Click to Search.... "
      Top             =   7230
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Ok"
      Enabled         =   0   'False
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
   Begin btButtonEx.ButtonEx cmdNew 
      Height          =   345
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Click to Add a New Appro..."
      Top             =   7230
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "New..."
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
   Begin btButtonEx.ButtonEx cmdReturn 
      Height          =   345
      Left            =   1290
      TabIndex        =   2
      Top             =   7230
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Open..."
      Enabled         =   0   'False
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
   Begin btButtonEx.ButtonEx ButtonEx4 
      Height          =   285
      Left            =   180
      TabIndex        =   12
      ToolTipText     =   "Click for a List of Appro's..."
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Name..."
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
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   825
      Left            =   1530
      TabIndex        =   16
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1455
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmAppro.frx":0084
   End
   Begin btButtonEx.ButtonEx cmdSale 
      Height          =   345
      Left            =   3990
      TabIndex        =   18
      Top             =   7230
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Convert to Sale"
      Enabled         =   0   'False
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
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   345
      Left            =   8070
      TabIndex        =   19
      ToolTipText     =   "Close appro's"
      Top             =   7230
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
      ShowFocus       =   0
   End
   Begin VB.Label lblAproNo 
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   630
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSForms.Image Image2 
      Height          =   295
      Left            =   1440
      Top             =   180
      Width           =   4515
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "7964;520"
   End
   Begin MSForms.Image Image10 
      Height          =   885
      Left            =   1440
      Top             =   1170
      Width           =   4515
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "7964;1561"
   End
   Begin VB.Label lblAcc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1590
      TabIndex        =   10
      Top             =   230
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   165
      Left            =   240
      TabIndex        =   9
      Top             =   1230
      Width           =   1140
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone No:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   570
      Width           =   1140
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID Number:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   900
      Width           =   1140
   End
   Begin MSForms.Image Image7 
      Height          =   295
      Left            =   1440
      Top             =   510
      Width           =   2535
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "4471;520"
   End
   Begin MSForms.Image Image4 
      Height          =   295
      Left            =   1440
      Top             =   840
      Width           =   2535
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "4471;520"
   End
   Begin MSForms.Image Image1 
      Height          =   2085
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   9255
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "16325;3678"
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   11
      Top             =   2250
      Width           =   7035
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   60
      Top             =   2190
      Width           =   9255
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "16325;661"
   End
   Begin MSForms.Image Image6 
      Height          =   8835
      Left            =   30
      Top             =   0
      Width           =   9390
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "16563;15584"
   End
End
Attribute VB_Name = "frmAppro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonEx1_Click()
    If cmdOk.Tag = "Finished" Then
    Unload Me
    End If
End Sub

Private Sub ButtonEx4_Click()
    grdMain.Cols = 3
    grdMain.TextMatrix(0, 0) = "Customer Name"
    grdMain.TextMatrix(0, 1) = "Telephone"
    grdMain.ColWidth(0) = grdMain.Width * 0.7
    grdMain.ColWidth(1) = grdMain.Width * 0.3
    grdMain.ColHidden(2) = True
    
    grdMain.Rows = 1
    ActiveReadServer "Select Cust_Name,Tel_No,Appro_No from Appros group by Cust_Name,Tel_No,Appro_No order by Cust_Name"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.Row = grdMain.Rows - 1
        grdMain.TextMatrix(grdMain.Row, 0) = rs.Fields("Cust_Name") & ""
        grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Tel_No") & ""
        grdMain.TextMatrix(grdMain.Row, 2) = rs.Fields("Appro_No") & ""
        rs.MoveNext
    Wend
    rs.Close
    If grdMain.Rows > 1 Then grdMain.Row = 1
    txtName.Locked = True
    txttel.Locked = True
    txtId.Locked = True
    txtAddress.Locked = True
    txtName.SetFocus
End Sub
Private Sub cmdCancel_Click()
grdMain.Cols = 3
    grdMain.Rows = 1
    grdMain.ColHidden(2) = True
    grdMain.TextMatrix(0, 0) = "Customer Name"
    grdMain.TextMatrix(0, 1) = "Telephone"
    grdMain.ColWidth(0) = grdMain.Width * 0.7
    grdMain.ColWidth(1) = grdMain.Width * 0.3
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    If frmAppro.Tag = "Not Now" Then
        frmAppro.Tag = ""
        Exit Sub
    End If
    grdMain.Rows = 1
    ActiveReadServer "Select Cust_Name,Tel_No,Appro_No from Appros group by Cust_Name,Tel_No,Appro_No order by Cust_Name"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.Row = grdMain.Rows - 1
        grdMain.TextMatrix(grdMain.Row, 0) = rs.Fields("Cust_Name") & ""
        grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Tel_No") & ""
        grdMain.TextMatrix(grdMain.Row, 2) = rs.Fields("Appro_No") & ""
        rs.MoveNext
    Wend
    rs.Close
    DoEvents
    Select Case grdMain.Cols
        Case 3
            If grdMain.TextMatrix(grdMain.Row, 2) <> "" Then
                ActiveReadServer "Select * from Appros where Appro_No = " & grdMain.TextMatrix(grdMain.Row, 2)
                If rs.RecordCount > 0 Then
                    txtName.Text = rs.Fields("Cust_Name") & ""
                    txttel.Text = rs.Fields("Tel_No") & ""
                    txtId.Text = rs.Fields("Id_No") & ""
                    txtAddress.Text = rs.Fields("Address") & ""
                    lblInfo.Caption = "Created on: " & Format(rs.Fields("Date_Time"), "DDD DD MMM YYYY")
                    cmdReturn.Caption = "Open..."
                    cmdReturn.Enabled = True
                End If
                rs.Close
            End If
    End Select
 cmdOk.Tag = "Finished"
End Sub
Private Sub cmdNew_Click()
    txtName.Text = ""
    txttel.Text = ""
    txtId.Text = ""
    txtAddress.Text = ""
    grdMain.Cols = 6
    grdMain.Rows = 1
    grdMain.TextMatrix(0, 0) = "Product Code"
    grdMain.TextMatrix(0, 1) = "Description"
    grdMain.TextMatrix(0, 2) = "Qty"
    grdMain.TextMatrix(0, 3) = "Cost"
    grdMain.TextMatrix(0, 4) = "Selling"
    'grdMain.ColHidden(5) = True
    grdMain.ColHidden(2) = False
    grdMain.ColHidden(3) = False
    grdMain.ColHidden(4) = False
    grdMain.ColWidth(0) = grdMain.Width * 0.15
    grdMain.ColWidth(1) = grdMain.Width * 0.47
    grdMain.ColWidth(2) = grdMain.Width * 0.1
    grdMain.ColWidth(3) = grdMain.Width * 0.1
    grdMain.ColWidth(4) = grdMain.Width * 0.1
    grdMain.ColWidth(4) = grdMain.Width * 0.1
    txtName.Locked = False
    txttel.Locked = False
    txtId.Locked = False
    txtAddress.Locked = False
    txtName.SetFocus
    cmdReturn.Enabled = False
    cmdReturn.Caption = "Open..."
End Sub

Private Sub cmdOk_Click()
    cmdOk.Tag = "Printing"
    Screen.MousePointer = 11
    If grdMain.TextMatrix(1, 5) = "" Then
        ActiveReadServer1 "Select isnull(max(Appro_No),0)+1 as Appro_No from Appros"
        lblAproNo.Caption = rs1.Fields("Appro_No")
        rs1.Close
    Else
        lblAproNo.Caption = grdMain.TextMatrix(1, 5)
    End If
    ActiveUpdateServer "Delete from Appros where Appro_No = " & lblAproNo.Caption
    DoEvents
    For i = 1 To grdMain.Rows - 1
        If grdMain.TextMatrix(i, 3) <> "" Then
            ActiveUpdateServer "Insert Into Appros (Cust_Name,Tel_No,ID_No,Address,Appro_No,Product_Code,Qty,Date_Time) values ('" & txtName.Text & "','" & txttel.Text & "','" & txtId.Text & "','" & txtAddress.Text & "','" & lblAproNo.Caption & "','" & grdMain.TextMatrix(i, 0) & "','" & grdMain.TextMatrix(i, 2) & "',Getdate())"
        End If
    Next i
    DoEvents
    Print_Appro 0
    ButtonEx4_Click
    Screen.MousePointer = 1
End Sub
Public Sub Print_Appro(Action)
On Error GoTo trap
With frmSlipDetails
    For b = 1 To 2
        PrintErr = 0
        Slip_Port = ""
        
        filenum = FreeFile
        Close #filenum
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
        If Action = 0 Then
            Print #filenum, "GOODS ON APPRO"
        Else
            Print #filenum, "APPRO GOODS RETURNED"
        End If
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
        If Slip_Printer_Type = 0 Then
            Print #filenum, String(40, "=")
        Else
            Print #filenum, String(33, "=")
        End If
        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        For i = 1 To grdMain.Rows - 1
            Print #filenum, Chr(27) & Chr(97) & Chr(48);
            Print #filenum, grdMain.TextMatrix(i, 0) & " - " & grdMain.TextMatrix(i, 1)
            Print #filenum, grdMain.TextMatrix(i, 2) & " @ R " & grdMain.TextMatrix(i, 4) & " = R " & Format(grdMain.TextMatrix(i, 4) * grdMain.TextMatrix(i, 2), "0.00")
        Next i
        Print #filenum, Chr(27) & Chr(50);
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
        Print #filenum, Chr(27) & Chr(97) & Chr(48);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        If Slip_Printer_Type = 0 Then
            Print #filenum, String(40, "=")
        Else
            Print #filenum, String(33, "=")
        End If
        Print #filenum, "    Customer: " & txtName.Text
        Print #filenum, "Telephone No: " & txttel.Text
        Print #filenum, "       ID No: " & txtId.Text
        Print #filenum, ""
        Print #filenum, ""
        Print #filenum, ""
        Print #filenum, ""
        Print #filenum, "Accepted: _____________________________"
        Print #filenum, ""
        Print #filenum, Chr(27) & Chr(97) & Chr(50);
        Print #filenum, Chr(27) & Chr(33) & Chr(0);
        If Slip_Printer_Type = 0 Then
            Print #filenum, String(40, "=")
        Else
            Print #filenum, String(33, "=")
        End If
        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
        Print #filenum, "Date: " & Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
        Print #filenum, "Appro No: " & Format(lblAproNo.Caption, "000000")
        For i = 1 To .grdFoot.Rows - 1
            If Trim(.grdFoot.TextMatrix(i, 0)) <> "" Then
                Select Case Trim(.grdFoot.TextMatrix(i, 1))
                    Case "Line Feeds"
                        Print #filenum, Chr(27) & Chr(100) & Chr(Val(.grdFoot.TextMatrix(i, 0)));
                    Case Else
                        Select Case .grdFoot.TextMatrix(i, 2)
                            Case "Left": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(48);
                            Case "Centre": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(49);
                            Case "Right": If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(97) & Chr(50);
                        End Select
                        If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(48);
                        If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(48);
                        Print #filenum, Chr(27) & Chr(33) & Chr(0);
                        Select Case Trim(.grdFoot.TextMatrix(i, 1))
                            Case ""
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case "Narrow Font"
                                If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case "Narrow Font (Dark)"
                                If Slip_Printer_Type = 0 Then Print #filenum, Chr(27) & Chr(77) & Chr(49);
                                If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case "Normal Font"
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case "Normal Font (Dark)"
                                If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case "Double Font"
                                Print #filenum, Chr(27) & Chr(33) & Chr(16);
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case "Double Font (Dark)"
                                Print #filenum, Chr(27) & Chr(33) & Chr(16);
                                If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case "Big Font"
                                Print #filenum, Chr(27) & Chr(33) & Chr(48);
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case "Big Font (Dark)"
                                Print #filenum, Chr(27) & Chr(33) & Chr(48);
                                If Slip_Printer_Type <> 2 Then Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                            Case Else
                                Print #filenum, .grdFoot.TextMatrix(i, 0)
                        End Select
                End Select
            End If
        Next i
        Print #filenum, Chr(29) & "V" & Chr(49);
        Close #filenum
    Next b
End With
 cmdOk.Tag = "Finished"
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
     cmdOk.Tag = "Finished"
    Close #filenum
    On Error GoTo 0
End Sub
Private Sub cmdReturn_Click()
    Select Case grdMain.Cols
        Case 3
            cmdReturn.Enabled = True
            cmdReturn.Caption = "Returned..."
            grdMain.Cols = 6
            grdMain.TextMatrix(0, 0) = "Product Code"
            grdMain.TextMatrix(0, 1) = "Description"
            grdMain.TextMatrix(0, 2) = "Qty"
            grdMain.TextMatrix(0, 3) = "Cost"
            grdMain.TextMatrix(0, 4) = "Selling"
            'grdMain.ColHidden(5) = True
            grdMain.ColHidden(2) = False
            grdMain.ColHidden(3) = False
            grdMain.ColHidden(4) = False
            grdMain.ColWidth(0) = grdMain.Width * 0.15
            grdMain.ColWidth(1) = grdMain.Width * 0.47
            grdMain.ColWidth(2) = grdMain.Width * 0.1
            grdMain.ColWidth(3) = grdMain.Width * 0.1
            grdMain.ColWidth(4) = grdMain.Width * 0.1
            grdMain.ColWidth(5) = grdMain.Width * 0.1
            grdMain.ColAlignment(2) = flexAlignRightCenter
            grdMain.ColAlignment(3) = flexAlignRightCenter
            grdMain.ColAlignment(4) = flexAlignRightCenter
            txtName.Locked = False
            txttel.Locked = False
            txtId.Locked = False
            txtAddress.Locked = False
            txtName.SetFocus
            ActiveReadServer "SELECT Appros.*, Products.Description,Products.Landed_Cost,Products.Selling_Price FROM Appros INNER JOIN Products ON Appros.Product_Code = Products.Product_Code Where Appros.Appro_No = " & grdMain.TextMatrix(grdMain.Row, 2)
            grdMain.Rows = 1
            If rs.RecordCount > 0 Then lblAproNo.Caption = rs.Fields("Appro_No")
            While Not rs.EOF
                grdMain.Rows = grdMain.Rows + 1
                grdMain.Row = grdMain.Rows - 1
                grdMain.TextMatrix(grdMain.Row, 0) = rs.Fields("Product_Code")
                grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Description")
                grdMain.TextMatrix(grdMain.Row, 2) = rs.Fields("Qty")
                grdMain.TextMatrix(grdMain.Row, 5) = rs.Fields("Appro_No")
                grdMain.TextMatrix(grdMain.Row, 3) = Format(rs.Fields("Landed_Cost"), "0.00")
                CCode = ""
                For i = 1 To Len(grdMain.TextMatrix(grdMain.Row, 3))
                    Select Case Mid(grdMain.TextMatrix(grdMain.Row, 3), i, 1)
                        Case "1": CCode = CCode & Cost_Code.One
                        Case "2": CCode = CCode & Cost_Code.Two
                        Case "3": CCode = CCode & Cost_Code.Three
                        Case "4": CCode = CCode & Cost_Code.Four
                        Case "5": CCode = CCode & Cost_Code.Five
                        Case "6": CCode = CCode & Cost_Code.Six
                        Case "7": CCode = CCode & Cost_Code.Seven
                        Case "8": CCode = CCode & Cost_Code.Eight
                        Case "9": CCode = CCode & Cost_Code.Nine
                        Case "0": CCode = CCode & Cost_Code.Ten
                        Case ".": CCode = CCode & "."
                    End Select
                Next i
                grdMain.TextMatrix(grdMain.Row, 3) = CCode
                grdMain.TextMatrix(grdMain.Row, 4) = Format(rs.Fields("Selling_Price"), "0.00")
                rs.MoveNext
            Wend
            rs.Close
        Case 6
             cmdOk.Tag = "Printing"
            Print_Appro 1
            ActiveReadServer "Delete FROM Appros Where Appros.Appro_No = " & lblAproNo.Caption
            ButtonEx4_Click
    End Select
End Sub

Private Sub Form_Activate()
    If frmAppro.Tag = "Not Now" Then
        frmAppro.Tag = ""
        Exit Sub
    End If
     cmdOk.Tag = "Finished"
    grdMain.Rows = 1
    ActiveReadServer "Select Cust_Name,Tel_No,Appro_No from Appros group by Cust_Name,Tel_No,Appro_No order by Cust_Name"
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.Row = grdMain.Rows - 1
        grdMain.TextMatrix(grdMain.Row, 0) = rs.Fields("Cust_Name") & ""
        grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Tel_No") & ""
        grdMain.TextMatrix(grdMain.Row, 2) = rs.Fields("Appro_No") & ""
        rs.MoveNext
    Wend
    rs.Close
    DoEvents
    Select Case grdMain.Cols
        Case 3
            If grdMain.TextMatrix(grdMain.Row, 2) <> "" Then
                ActiveReadServer "Select * from Appros where Appro_No = " & grdMain.TextMatrix(grdMain.Row, 2)
                If rs.RecordCount > 0 Then
                    txtName.Text = rs.Fields("Cust_Name") & ""
                    txttel.Text = rs.Fields("Tel_No") & ""
                    txtId.Text = rs.Fields("Id_No") & ""
                    txtAddress.Text = rs.Fields("Address") & ""
                    lblInfo.Caption = "Created on: " & Format(rs.Fields("Date_Time"), "DDD DD MMM YYYY")
                    cmdReturn.Caption = "Open..."
                    cmdReturn.Enabled = True
                End If
                rs.Close
            End If
    End Select
End Sub
Private Sub Form_Load()
    grdMain.Cols = 3
    grdMain.Rows = 1
    grdMain.ColHidden(2) = True
    grdMain.TextMatrix(0, 0) = "Customer Name"
    grdMain.TextMatrix(0, 1) = "Telephone"
    grdMain.ColWidth(0) = grdMain.Width * 0.7
    grdMain.ColWidth(1) = grdMain.Width * 0.3
    grdMain.ColAlignment(1) = flexAlignLeftCenter
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmSearch
End Sub
Private Sub grdMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
        Case 0
            ActiveReadServer "Select Description,Landed_Cost,Selling_Price from Products where Product_Code = '" & grdMain.TextMatrix(Row, 0) & "' and Sales_Item = 1"
            If rs.RecordCount > 0 Then
                grdMain.TextMatrix(Row, 1) = rs.Fields("Description") & ""
                grdMain.TextMatrix(Row, 2) = "0"
                grdMain.TextMatrix(Row, 3) = Format(rs.Fields("Landed_Cost"), "0.00")
                CCode = ""
                For i = 1 To Len(grdMain.TextMatrix(Row, 3))
                    Select Case Mid(grdMain.TextMatrix(Row, 3), i, 1)
                        Case "1": CCode = CCode & Cost_Code.One
                        Case "2": CCode = CCode & Cost_Code.Two
                        Case "3": CCode = CCode & Cost_Code.Three
                        Case "4": CCode = CCode & Cost_Code.Four
                        Case "5": CCode = CCode & Cost_Code.Five
                        Case "6": CCode = CCode & Cost_Code.Six
                        Case "7": CCode = CCode & Cost_Code.Seven
                        Case "8": CCode = CCode & Cost_Code.Eight
                        Case "9": CCode = CCode & Cost_Code.Nine
                        Case "0": CCode = CCode & Cost_Code.Ten
                        Case ".": CCode = CCode & "."
                    End Select
                Next i
                grdMain.TextMatrix(Row, 3) = CCode
                If CCode = "." Then
                    grdMain.TextMatrix(Row, 3) = Format(rs.Fields("Landed_Cost"), "0.00")
                End If
                grdMain.TextMatrix(Row, 4) = Format(rs.Fields("Selling_Price"), "0.00")
            Else
                grdMain.TextMatrix(Row, 0) = ""
            End If
            rs.Close
    End Select
End Sub
Private Sub grdMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Select Case grdMain.Cols
        Case 3
            If grdMain.TextMatrix(NewRow, 2) <> "" Then
                ActiveReadServer "Select * from Appros where Appro_No = " & grdMain.TextMatrix(NewRow, 2)
                If rs.RecordCount > 0 Then
                    txtName.Text = rs.Fields("Cust_Name") & ""
                    txttel.Text = rs.Fields("Tel_No") & ""
                    txtId.Text = rs.Fields("Id_No") & ""
                    txtAddress.Text = rs.Fields("Address") & ""
                    lblInfo.Caption = "Created on: " & Format(rs.Fields("Date_Time"), "DDD DD MMM YYYY")
                    cmdReturn.Caption = "Open..."
                    cmdReturn.Enabled = True
                End If
                rs.Close
            End If
    End Select
End Sub
Private Sub grdMain_AfterSort(ByVal Col As Long, Order As Integer)
    Select Case grdMain.Cols
        Case 3
            If grdMain.TextMatrix(NewRow, 2) <> "" Then
                ActiveReadServer "Select * from Appros where Appro_No = " & grdMain.TextMatrix(NewRow, 2)
                If rs.RecordCount > 0 Then
                    txtName.Text = rs.Fields("Cust_Name") & ""
                    txttel.Text = rs.Fields("Tel_No") & ""
                    txtId.Text = rs.Fields("Id_No") & ""
                    txtAddress.Text = rs.Fields("Address") & ""
                    lblInfo.Caption = "Created on: " & Format(rs.Fields("Date_Time"), "DDD DD MMM YYYY")
                    cmdReturn.Caption = "Open..."
                    cmdReturn.Enabled = True
                End If
                rs.Close
            End If
    End Select
End Sub

Private Sub grdMain_CellChanged(ByVal Row As Long, ByVal Col As Long)
    CheckforSave
End Sub

Private Sub grdMain_EnterCell()
    If grdMain.Cols = 6 Then
        Select Case grdMain.Col
            Case 0
                grdMain.Editable = flexEDKbdMouse
            Case 1
                grdMain.Editable = flexEDNone
            Case 2
                grdMain.Editable = flexEDKbdMouse
            Case Else
                grdMain.Editable = flexEDNone
        End Select
    Else
        grdMain.Editable = flexEDNone
    End If
End Sub
Private Sub grdMain_GotFocus()
    Select Case grdMain.Cols
        Case 6
            grdMain.SelectionMode = flexSelectionFree
            If grdMain.Rows = 1 Then
                grdMain.Rows = 2
                grdMain.Row = 1
                grdMain.Col = 0
                grdMain.Editable = flexEDKbdMouse
            End If
        Case 3
            grdMain.SelectionMode = flexSelectionByRow
            If grdMain.Rows = 1 Then txtName.SetFocus
    End Select
End Sub
Private Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            If grdMain.Row = 1 Then
                txtAddress.SetFocus
                If grdMain.TextMatrix(grdMain.Rows - 1, 0) = "" Then
                    grdMain.RemoveItem grdMain.Rows - 1
                End If
            End If
        Case 40
            If grdMain.Cols = 6 Then
                If grdMain.Row = grdMain.Rows - 1 Then
                    If grdMain.TextMatrix(grdMain.Rows - 1, 0) = "" Then
                        grdMain.RemoveItem grdMain.Rows - 1
                        txtName.SetFocus
                    Else
                        grdMain.Rows = grdMain.Rows + 1
                    End If
                End If
            End If
        Case 46
            grdMain.RemoveItem (grdMain.Row)
            If grdMain.Rows = 1 Then
                grdMain.Rows = 2
                grdMain.Row = 1
                grdMain.Col = 0
            End If
        Case 13
            If grdMain.Col = 0 Then
                Screen.MousePointer = 11
                Load frmSearch
                frmSearch.Tag = "Appros"
                DoEvents
                Screen.MousePointer = 0
                frmAppro.Tag = "Not Now"
                frmSearch.Show vbModal
                Select Case frmSearch.Tag
                    Case ""
                        grdMain.TextMatrix(grdMain.Row, 0) = ""
                    Case Else
                        Product_Code = Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1))
                        ActiveReadServer "Select Description,Landed_Cost,Selling_Price from Products where Product_Code = '" & Product_Code & "' and Sales_Item = 1"
                        If rs.RecordCount > 0 Then
                            grdMain.TextMatrix(grdMain.Row, 0) = Product_Code
                            grdMain.TextMatrix(grdMain.Row, 1) = rs.Fields("Description") & ""
                            grdMain.TextMatrix(grdMain.Row, 2) = "0"
                            grdMain.TextMatrix(grdMain.Row, 3) = Format(rs.Fields("Landed_Cost"), "0.00")
                            CCode = ""
                            For i = 1 To Len(grdMain.TextMatrix(grdMain.Row, 3))
                                Select Case Mid(grdMain.TextMatrix(grdMain.Row, 3), i, 1)
                                    Case "1": CCode = CCode & Cost_Code.One
                                    Case "2": CCode = CCode & Cost_Code.Two
                                    Case "3": CCode = CCode & Cost_Code.Three
                                    Case "4": CCode = CCode & Cost_Code.Four
                                    Case "5": CCode = CCode & Cost_Code.Five
                                    Case "6": CCode = CCode & Cost_Code.Six
                                    Case "7": CCode = CCode & Cost_Code.Seven
                                    Case "8": CCode = CCode & Cost_Code.Eight
                                    Case "9": CCode = CCode & Cost_Code.Nine
                                    Case "0": CCode = CCode & Cost_Code.Ten
                                    Case ".": CCode = CCode & "."
                                End Select
                            Next i
                            grdMain.TextMatrix(grdMain.Row, 3) = CCode
                            If CCode = "." Then
                                grdMain.TextMatrix(grdMain.Row, 3) = Format(rs.Fields("Landed_Cost"), "0.00")
                            End If
                            grdMain.TextMatrix(grdMain.Row, 4) = Format(rs.Fields("Selling_Price"), "0.00")
                        Else
                            grdMain.TextMatrix(grdMain.Row, 0) = ""
                        End If
                        rs.Close
                        frmSearch.Tag = ""
                End Select
            End If
        Case 45, 48 To 57, 65 To 90, 96 To 105, 109, 110, 189
            Select Case grdMain.Col
                Case 1, 3, 4
                Case Else
                    If grdMain.Cols = 6 Then grdMain.EditCell
            End Select
        Case 37 'left
            If grdMain.Col = 0 Then
                KeyCode = 0
                grdMain.Col = 2
            End If
    End Select
    CheckforSave
End Sub
Private Sub grdMain_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case grdMain.Col
        Case 0
            Select Case KeyAscii
                Case 8, 13, 48 To 57
                Case 39
                    KeyAscii = 0
                Case 65 To 90
                Case 97 To 122
                    KeyAscii = KeyAscii - 32
                Case Else
                    KeyAscii = 0
            End Select
        Case 2
            Select Case KeyAscii
                Case 8, 13, 48 To 57
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub





Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            If txtAddress.GetLineFromChar(txtAddress.SelStart) = 0 Then
                KeyCode = 0
                txtId.SetFocus
            End If
        Case 13, 40
            If KeyCode = 40 And InStr(Mid(txtAddress.Text, txtAddress.SelStart + 1), Chr(13)) = 0 Then
                KeyCode = 0
                grdMain.SetFocus
            End If
            If txtAddress.GetLineFromChar(txtAddress.SelStart) = 4 Then
                KeyCode = 0
                grdMain.SetFocus
            End If
    End Select
End Sub
Private Sub txtId_Change()
    CheckforSave
End Sub
Private Sub txtID_GotFocus()
    txtId.SelStart = 0
    txtId.SelLength = Len(txtId.Text)
End Sub
Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            KeyCode = 0
            txttel.SetFocus
        Case 13, 40
            KeyCode = 0
            txtAddress.SetFocus
    End Select
End Sub
Private Sub txtID_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 32
        Case 39
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 45, 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtName_Change()
    CheckforSave
End Sub
Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            KeyCode = 0
            grdMain.SetFocus
        Case 13, 40
            KeyCode = 0
            txttel.SetFocus
    End Select
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtName_LostFocus()
    On Error Resume Next
    txtName.Text = UCase(Left(txtName.Text, 1)) & Mid(txtName.Text, 2)
    On Error GoTo 0
End Sub
Private Sub CheckforSave()
    cmdOk.Enabled = True
    cmdSale.Enabled = True
    If txtName.Text = "" Then cmdOk.Enabled = False
    If txtId.Text = "" Then cmdOk.Enabled = False
    If txttel.Text = "" Then cmdOk.Enabled = False
    If grdMain.Rows = 1 Then cmdOk.Enabled = False
    If grdMain.TextMatrix(grdMain.Rows - 1, 0) = "" Then
        If grdMain.Rows = 2 Then
            cmdOk.Enabled = False
        End If
    End If
    If grdMain.Cols = 3 Then cmdOk.Enabled = False
    If cmdOk.Enabled = False Then
        cmdSale.Enabled = False
    End If
End Sub
Private Sub txtTel_Change()
    CheckforSave
End Sub
Private Sub txtTel_GotFocus()
    txttel.SelStart = 0
    txttel.SelLength = Len(txttel.Text)
End Sub
Private Sub txtTel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            KeyCode = 0
            txtName.SetFocus
        Case 13, 40
            KeyCode = 0
            txtId.SetFocus
    End Select
End Sub
Private Sub txtTel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
        Case 32
        Case 39
            KeyAscii = 0
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case 45, 48 To 57, 65 To 90
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtAddress_Change()
    On Error Resume Next
    Rows = 0
    For i = 1 To Len(txtAddress.Text)
        If Asc(Mid(txtAddress.Text, i, 1)) = 13 Then
            Rows = Rows + 1
            If Rows = 5 Then
                grdMain.SetFocus
                Exit For
            End If
        End If
    Next i
    CheckforSave
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

