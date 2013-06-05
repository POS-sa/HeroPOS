VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSlipDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Slip Details"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "frmSlipDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   9075
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Index           =   0
      Left            =   9180
      TabIndex        =   2
      Top             =   8640
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
   Begin VSFlex8Ctl.VSFlexGrid grdHead 
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   10245
      _cx             =   18071
      _cy             =   6853
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   13
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid grdFoot 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   4890
      Width           =   10245
      _cx             =   18071
      _cy             =   6376
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Index           =   1
      Left            =   7890
      TabIndex        =   3
      Top             =   8640
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
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   8640
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Test"
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
   Begin VB.Label Label2 
      Caption         =   "Footer"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   5
      Top             =   4560
      Width           =   1755
   End
   Begin MSForms.Label Label1 
      Height          =   285
      Left            =   210
      TabIndex        =   4
      Top             =   180
      Width           =   8805
      Caption         =   "Header"
      Size            =   "15531;503"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image4 
      Height          =   3645
      Left            =   90
      Top             =   4890
      Width           =   10305
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "18177;6429"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image1 
      Height          =   405
      Left            =   90
      Top             =   4500
      Width           =   10305
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "18177;714"
   End
   Begin MSForms.Image Image3 
      Height          =   3915
      Left            =   90
      Top             =   510
      Width           =   10305
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "18177;6906"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image2 
      Height          =   405
      Left            =   90
      Top             =   120
      Width           =   10305
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "18177;714"
   End
End
Attribute VB_Name = "frmSlipDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEnd_Click(Index As Integer)
    On Error GoTo trap
    PrintErr = 0
    Slip_Port = ""
    Select Case cmdEnd(Index).Caption
        Case "Test"
            ActiveReadServer "Delete from Printer_Header"
            For i = 1 To grdHead.Rows - 1
                ActiveUpdateServer "Insert Into Printer_Header (Description,Style,Alignment) values ('" & Replace(grdHead.TextMatrix(i, 0), "'", "`") & "','" & grdHead.TextMatrix(i, 1) & "','" & grdHead.TextMatrix(i, 2) & "')"
            Next i
            ActiveReadServer "Delete from Printer_Footer"
            For i = 1 To grdFoot.Rows - 1
                ActiveUpdateServer "Insert Into Printer_Footer (Description,Style,Alignment) values ('" & Replace(grdFoot.TextMatrix(i, 0), "'", "`") & "','" & grdFoot.TextMatrix(i, 1) & "','" & grdFoot.TextMatrix(i, 2) & "')"
            Next i
            DoEvents
            filenum = FreeFile
            If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
                If InStr(Trim(Slip_Printer), "\\") = 0 Then
                    If Slip_Port = "" Then
                        Open "\\" & Comp_Name & "\" & Slip_Printer For Output As filenum
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
            Print #filenum, Chr(27) & Chr(69) & Chr(1);
            For i = 1 To grdHead.Rows - 1
                If Trim(grdHead.TextMatrix(i, 0)) <> "" Then
                    Select Case Trim(grdHead.TextMatrix(i, 1))
                        Case "Line Feeds"
                            Print #filenum, Chr(27) & Chr(100) & Chr(Val(grdHead.TextMatrix(i, 0)));
                        Case "Full Cut"
                            Print #filenum, Chr(29) & "V" & Chr(49);
                        Case "Partial Cut"
                            Print #filenum, Chr(29) & "V" & Chr(50);
                        Case Else
                            Select Case grdHead.TextMatrix(i, 2)
                                Case "Left": Print #filenum, Chr(27) & Chr(97) & Chr(48);
                                Case "Centre": Print #filenum, Chr(27) & Chr(97) & Chr(49);
                                Case "Right": Print #filenum, Chr(27) & Chr(97) & Chr(50);
                            End Select
                            Print #filenum, Chr(27) & Chr(69) & Chr(48);
                            Print #filenum, Chr(27) & Chr(77) & Chr(48);
                            Print #filenum, Chr(27) & Chr(33) & Chr(0);
                            Select Case Trim(grdHead.TextMatrix(i, 1))
                                Case ""
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case "Narrow Font"
                                    Print #filenum, Chr(27) & Chr(77) & Chr(49);
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case "Narrow Font (Dark)"
                                    Print #filenum, Chr(27) & Chr(77) & Chr(49);
                                    Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case "Normal Font"
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case "Normal Font (Dark)"
                                    Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case "Double Font"
                                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case "Double Font (Dark)"
                                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                                    Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case "Big Font"
                                    Print #filenum, Chr(27) & Chr(33) & Chr(48);
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case "Big Font (Dark)"
                                    Print #filenum, Chr(27) & Chr(33) & Chr(48);
                                    Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                                Case Else
                                    Print #filenum, grdHead.TextMatrix(i, 0)
                            End Select
                    End Select
                End If
            Next i
            Print #filenum, ""
            Print #filenum, ""
            Print #filenum, Chr(27) & Chr(33) & Chr(16);
            Print #filenum, Chr(27) & Chr(69) & Chr(49);
            Print #filenum, Chr(27) & Chr(97) & Chr(49);
            Print #filenum, "SALE BODY"
            Print #filenum, Chr(27) & Chr(69) & Chr(48);
            Print #filenum, Chr(27) & Chr(77) & Chr(48);
            Print #filenum, Chr(27) & Chr(33) & Chr(0);
            Print #filenum, ""
            Print #filenum, ""
            For i = 1 To grdFoot.Rows - 1
                If Trim(grdFoot.TextMatrix(i, 0)) <> "" Then
                    Select Case Trim(grdFoot.TextMatrix(i, 1))
                        Case "Line Feeds"
                            Print #filenum, Chr(27) & Chr(100) & Chr(Val(grdFoot.TextMatrix(i, 0)));
                        Case Else
                            Select Case grdFoot.TextMatrix(i, 2)
                                Case "Left": Print #filenum, Chr(27) & Chr(97) & Chr(48);
                                Case "Centre": Print #filenum, Chr(27) & Chr(97) & Chr(49);
                                Case "Right": Print #filenum, Chr(27) & Chr(97) & Chr(50);
                            End Select
                            Print #filenum, Chr(27) & Chr(69) & Chr(48);
                            Print #filenum, Chr(27) & Chr(77) & Chr(48);
                            Print #filenum, Chr(27) & Chr(33) & Chr(0);
                            Select Case Trim(grdFoot.TextMatrix(i, 1))
                                Case ""
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case "Narrow Font"
                                    Print #filenum, Chr(27) & Chr(77) & Chr(49);
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case "Narrow Font (Dark)"
                                    Print #filenum, Chr(27) & Chr(77) & Chr(49);
                                    Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case "Normal Font"
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case "Normal Font (Dark)"
                                    Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case "Double Font"
                                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case "Double Font (Dark)"
                                    Print #filenum, Chr(27) & Chr(33) & Chr(16);
                                    Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case "Big Font"
                                    Print #filenum, Chr(27) & Chr(33) & Chr(48);
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case "Big Font (Dark)"
                                    Print #filenum, Chr(27) & Chr(33) & Chr(48);
                                    Print #filenum, Chr(27) & Chr(69) & Chr(49);
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                                Case Else
                                    Print #filenum, grdFoot.TextMatrix(i, 0)
                            End Select
                    End Select
                End If
            Next i
            Print #filenum, Chr(27) & Chr(33) & Chr(1);
            Print #filenum, Chr(27) & Chr(97) & Chr(48);
            Close #filenum
        Case "Cancel"
            Unload Me
        Case "Ok"
            ActiveReadServer "Delete from Printer_Header"
            For i = 1 To grdHead.Rows - 1
                ActiveUpdateServer "Insert Into Printer_Header (Description,Style,Alignment) values ('" & Replace(grdHead.TextMatrix(i, 0), "'", "`") & "','" & grdHead.TextMatrix(i, 1) & "','" & grdHead.TextMatrix(i, 2) & "')"
            Next i
            ActiveReadServer "Delete from Printer_Footer"
            For i = 1 To grdFoot.Rows - 1
                ActiveUpdateServer "Insert Into Printer_Footer (Description,Style,Alignment) values ('" & Replace(grdFoot.TextMatrix(i, 0), "'", "`") & "','" & grdFoot.TextMatrix(i, 1) & "','" & grdFoot.TextMatrix(i, 2) & "')"
            Next i
            MsgBox "Printer Details Updated.", vbInformation, "HeroPOS Information Message"
            
            Unload Me
    End Select
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
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    grdHead.SetFocus
    grdHead.Row = 1
    grdHead_EnterCell
End Sub

Private Sub Form_Load()
    grdHead.TextMatrix(0, 0) = "Line Detail"
    grdFoot.TextMatrix(0, 1) = "Style"
    grdHead.TextMatrix(0, 0) = "Line Detail"
    grdHead.TextMatrix(0, 1) = "Style"
    grdHead.TextMatrix(0, 2) = "Alignment"
    grdFoot.TextMatrix(0, 2) = "Alignment"
    grdHead.ColAlignment(0) = flexAlignLeftCenter
    grdHead.ColAlignment(1) = flexAlignLeftCenter
    grdFoot.ColAlignment(0) = flexAlignLeftCenter
    grdFoot.ColAlignment(1) = flexAlignLeftCenter
    grdHead.ColWidth(0) = grdHead.Width * 0.6
    grdHead.ColWidth(1) = grdHead.Width * 0.3
    grdHead.ColWidth(2) = grdHead.Width * 0.1
    grdFoot.ColWidth(0) = grdHead.Width * 0.6
    grdFoot.ColWidth(1) = grdHead.Width * 0.3
    grdFoot.ColWidth(2) = grdHead.Width * 0.1
    ActiveReadServer "Select * from Printer_Header order by Line_No"
    i = 0
    While Not rs.EOF
        i = i + 1
        grdHead.TextMatrix(i, 0) = Replace(Trim(rs.Fields("Description") & ""), "`", "'")
        grdHead.TextMatrix(i, 1) = Trim(rs.Fields("Style") & "")
        grdHead.TextMatrix(i, 2) = Trim(rs.Fields("Alignment") & "")
        If grdHead.TextMatrix(i, 0) <> "" And grdHead.TextMatrix(i, 2) = "" Then
            grdHead.TextMatrix(i, 2) = "Left"
        End If
        If grdHead.TextMatrix(i, 0) = "" Then
            grdHead.TextMatrix(i, 1) = ""
            grdHead.TextMatrix(i, 2) = ""
        End If
        rs.MoveNext
    Wend
    rs.Close
    ActiveReadServer "Select * from Printer_Footer order by Line_No"
    i = 0
    While Not rs.EOF
        i = i + 1
        grdFoot.TextMatrix(i, 0) = Replace(Trim(rs.Fields("Description") & ""), "`", "'")
        grdFoot.TextMatrix(i, 1) = Trim(rs.Fields("Style") & "")
        grdFoot.TextMatrix(i, 2) = Trim(rs.Fields("Alignment") & "")
        If grdFoot.TextMatrix(i, 0) <> "" And grdFoot.TextMatrix(i, 2) = "" Then
            grdFoot.TextMatrix(i, 2) = "Left"
        End If
        If grdFoot.TextMatrix(i, 0) = "" Then
            If grdFoot.TextMatrix(i, 1) <> "Partial Cut" Then
                If grdFoot.TextMatrix(i, 1) <> "Full Cut" Then
                    grdFoot.TextMatrix(i, 1) = ""
                End If
            End If
            grdFoot.TextMatrix(i, 2) = ""
        End If
        rs.MoveNext
    Wend
    
    rs.Close
End Sub

Private Sub grdFoot_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then
    If grdFoot.TextMatrix(Row, Col) <> "" And grdFoot.TextMatrix(Row, 2) = "" Then
        grdFoot.TextMatrix(Row, 2) = "Left"
    End If
    If grdFoot.TextMatrix(Row, Col) = "" Then
        grdFoot.TextMatrix(Row, 2) = ""
    End If
End If
End Sub
Private Sub grdHead_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    If grdHead.TextMatrix(Row, Col) <> "" And grdHead.TextMatrix(Row, 2) = "" Then
        grdHead.TextMatrix(Row, 2) = "Left"
    End If
    If grdHead.TextMatrix(Row, Col) = "" Then
        grdHead.TextMatrix(Row, 2) = ""
    End If
    
End If
End Sub

Private Sub grdHead_EnterCell()
    Select Case grdHead.Col
        Case 0
            grdHead.Editable = flexEDKbdMouse
            grdHead.ShowComboButton = flexSBFocus
        Case 1
            grdHead.ColComboList(1) = "Narrow Font|Narrow Font (Dark)|Normal Font|Normal Font (Dark)|Double Font|Double Font (Dark)|Big Font|Big Font (Dark)|Line Feeds"
        Case 2
            grdHead.ColComboList(2) = "Left|Centre|Right"
    End Select
End Sub
Private Sub grdFoot_EnterCell()
    Select Case grdFoot.Col
        Case 0
            grdFoot.Editable = flexEDKbdMouse
            grdFoot.ShowComboButton = flexSBFocus
        Case 1
            grdFoot.ColComboList(1) = "Narrow Font|Narrow Font (Dark)|Normal Font|Normal Font (Dark)|Double Font|Double Font (Dark)|Big Font|Big Font (Dark)|Line Feeds|Full Cut|Partial Cut"
        Case 2
            grdFoot.ColComboList(2) = "Left|Centre|Right"
    End Select
End Sub

Private Sub grdHead_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46
            grdHead.TextMatrix(grdHead.Row, 0) = ""
            grdHead.TextMatrix(grdHead.Row, 1) = ""
            grdHead.TextMatrix(grdHead.Row, 2) = ""
    End Select
End Sub
Private Sub grdfoot_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46
            grdFoot.TextMatrix(grdFoot.Row, 0) = ""
            grdFoot.TextMatrix(grdFoot.Row, 1) = ""
            grdFoot.TextMatrix(grdFoot.Row, 2) = ""
    End Select
End Sub

