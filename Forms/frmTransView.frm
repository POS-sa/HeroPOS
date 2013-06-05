VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTransView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transaction View..."
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   FillColor       =   &H00C0FFC0&
   Icon            =   "frmTransView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   375
      Left            =   6930
      TabIndex        =   3
      Top             =   8940
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
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
   Begin VSFlex8Ctl.VSFlexGrid grdMain 
      Height          =   6735
      Left            =   60
      TabIndex        =   1
      Top             =   1920
      Width           =   8235
      _cx             =   14526
      _cy             =   11880
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   15523287
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
      WallPaper       =   "frmTransView.frx":000C
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin btButtonEx.ButtonEx cmdPrint 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   8940
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Reprint"
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
   Begin btButtonEx.ButtonEx cmdTender 
      Height          =   495
      Left            =   210
      TabIndex        =   19
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Change Tender"
      CaptionOffsetY  =   1
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
   End
   Begin VB.Label txtTender 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   9075
      Width           =   3165
   End
   Begin VB.Label txtOwner 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   1680
      TabIndex        =   17
      Top             =   8745
      Width           =   3165
   End
   Begin MSForms.Label Label5 
      Height          =   255
      Left            =   -270
      TabIndex        =   16
      Top             =   8730
      Width           =   1785
      ForeColor       =   7555868
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Originated By:"
      Size            =   "3149;450"
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   -270
      TabIndex        =   15
      Top             =   9075
      Width           =   1785
      ForeColor       =   7555868
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Amount Tendered:"
      Size            =   "3149;450"
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Invoice No"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   210
      TabIndex        =   14
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1650
      TabIndex        =   13
      Top             =   630
      Width           =   3585
   End
   Begin VB.Label lblAcc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1650
      TabIndex        =   12
      Top             =   930
      Width           =   3585
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1650
      TabIndex        =   11
      Top             =   330
      Width           =   3585
   End
   Begin MSForms.Image Image4 
      Height          =   1215
      Left            =   1410
      Top             =   180
      Width           =   4215
      BackColor       =   16777215
      Size            =   "7435;2143"
   End
   Begin VB.Label lblDocNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   570
      Width           =   1095
   End
   Begin VB.Label lblTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   1500
      Width           =   7785
   End
   Begin MSForms.Image Image3 
      Height          =   315
      Left            =   150
      Top             =   1470
      Width           =   8085
      BackColor       =   16777215
      Size            =   "14261;556"
   End
   Begin MSForms.Label Label4 
      Height          =   315
      Left            =   5610
      TabIndex        =   8
      Top             =   1050
      Width           =   855
      ForeColor       =   7555868
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Total:"
      Size            =   "1508;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label3 
      Height          =   315
      Left            =   5610
      TabIndex        =   7
      Top             =   600
      Width           =   855
      ForeColor       =   7555868
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Vat:"
      Size            =   "1508;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblTotal 
      Height          =   360
      Left            =   6510
      TabIndex        =   6
      Top             =   1020
      Width           =   1725
      ForeColor       =   7555868
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "3043;635"
      BorderStyle     =   1
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblVat 
      Height          =   360
      Left            =   6510
      TabIndex        =   5
      Top             =   600
      Width           =   1725
      ForeColor       =   7555868
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "3043;635"
      BorderStyle     =   1
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblSubTotal 
      Height          =   360
      Left            =   6510
      TabIndex        =   0
      Top             =   180
      Width           =   1725
      ForeColor       =   7555868
      BackColor       =   16777215
      Caption         =   "0.00"
      Size            =   "3043;635"
      BorderStyle     =   1
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCash 
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   210
      Width           =   855
      ForeColor       =   7555868
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Subtotal:"
      Size            =   "1508;556"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image2 
      Height          =   1215
      Left            =   150
      Top             =   180
      Width           =   1215
      BackColor       =   16777215
      Size            =   "2143;2143"
   End
   Begin MSForms.Image Image1 
      Height          =   1785
      Left            =   60
      Top             =   90
      Width           =   8250
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "14552;3149"
   End
   Begin MSForms.Image Image5 
      Height          =   285
      Left            =   1560
      Top             =   8730
      Width           =   3465
      BackColor       =   16777215
      Size            =   "6112;503"
   End
   Begin MSForms.Image Image6 
      Height          =   285
      Left            =   1560
      Top             =   9060
      Width           =   3465
      BackColor       =   16777215
      Size            =   "6112;503"
   End
End
Attribute VB_Name = "frmTransView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    Panel_No = -1
    Mystring = grdMain.TextMatrix(grdMain.Rows - 1, 1)
    If Mystring = "Discount Value" Then
        TillData.TotDiscount = grdMain.TextMatrix(grdMain.Rows - 1, 2)
        Reprintdiscount = 0
        Thediscounttotal = 0
        Select Case grdMain.TextMatrix(grdMain.Rows - 2, 1)
            Case "Cash Tendered"
                PrintSlip "Cash"
            Case "Card Tendered"
                PrintSlip "Card"
            Case "Charge Tendered"
                PrintSlip "Charge"
            Case "Voucher Tendered"
                PrintSlip "Voucher"
        End Select
    
    
    End If
    Mystring = grdMain.TextMatrix(grdMain.Rows - 1, 1)
    If Mystring <> "Discount Discount Value" Then
        Select Case grdMain.TextMatrix(grdMain.Rows - 1, 1)
            Case "Cash Tendered"
                PrintSlip "Cash"
            Case "Card Tendered"
                PrintSlip "Card"
            Case "Charge Tendered"
                PrintSlip "Charge"
            Case "Voucher Tendered"
                PrintSlip "Voucher"
        End Select
    End If
    
    
    
    
    Panel_No = 0
    Unload Me
End Sub

Private Sub cmdTender_Click()
    frmTender.Show vbModal
    Form_Activate
End Sub
Private Sub Form_Activate()
    If frmTransView.Tag = "" Then
        LoadOldTable frmReports.grdMain.TextMatrix(frmReports.grdMain.Row, 3)
    Else
        LoadOldTable frmTransView.Tag
    End If
    If UserRecord.uType = 9 Or UserRecord.uType = 0 Then
        cmdTender.Visible = True
    Else
        cmdTender.Visible = False
    End If
End Sub
Public Sub LoadOldTable(Table_No)
    On Error Resume Next
    ActiveReadServer "Select * from Transaction_View where Function_Key <> 14 and Invoice_No= " & Val(Table_No) & " order by Line_No"
    grdMain.Rows = 1
    grdMain.ColHidden(14) = True
    If rs.RecordCount > 0 Then
        ActiveReadServer2 "Select First_Name,Last_Name from Users where User_No = " & Val(rs.Fields("User_No") & "")
        If rs2.RecordCount > 0 Then
            lblUser.Caption = "User: " & rs.Fields("User_No") & " - " & rs2.Fields("First_name") & " " & rs2.Fields("Last_Name")
        Else
            lblUser.Caption = "User: " & rs.Fields("User_No") & " - Deleted User"
        End If
        lblDate.Caption = "Dated: " & Format(rs.Fields("Date_Time"), "DDD DD MMM YYYY HH:mm:SS")
        lblDocNo = String(7 - Len(rs.Fields("Invoice_No")), "0") & rs.Fields("Invoice_No")

        rs2.Close
        SCharge = 0
        TillData.Tipp = 0
        txtTender.Caption = ""
        While Not rs.EOF
top:
            If rs.Fields("Function_Key") <> 14 Then
                grdMain.Rows = grdMain.Rows + 1
                grdMain.TextMatrix(grdMain.Rows - 1, 0) = Trim(rs.Fields("Qty") & "")
                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Short_Description") & ""
                If rs.Fields("Short_Description") & "" = "" And rs.Fields("Function_Key") = 7 Then
                    grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Extra") & ""
                    If Val(rs.Fields("Qty") & "") <> 0 And rs.Fields("Extra") & "" = "" Then
                        grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Deleted Product"
                        grdMain.RemoveItem grdMain.Rows - 1
                        rs.MoveNext
                        GoTo top
                    End If
                End If
                If rs.Fields("Short_Description") & "" = "" And rs.Fields("Function_Key") = 16 Then
                    grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Service Charge"
                End If
                If rs.Fields("Function_Key") = 6 Then
                    grdMain.TextMatrix(grdMain.Rows - 1, 1) = "No Sale"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0FFC0
                End If
                
                If Val(rs.Fields("Line_Total") & "") <> 0 Then
                    grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(Val(rs.Fields("Line_Total") & ""), "0.00")
                    Select Case Val(rs.Fields("Function_Key") & "")
                        Case 9: txtTender.Caption = Format(Val(rs.Fields("Ave_Cost") & ""), "0.00")
                        Case 10: txtTender.Caption = Format(Val(rs.Fields("Ave_Cost") & ""), "0.00")
                        Case 11: txtTender.Caption = Format(Val(rs.Fields("Ave_Cost") & ""), "0.00")
                        Case 12: txtTender.Caption = Format(Val(rs.Fields("Ave_Cost") & ""), "0.00")
                    End Select
                End If
                grdMain.TextMatrix(grdMain.Rows - 1, 4) = Val(rs.Fields("Ave_Cost") & "")
                grdMain.TextMatrix(grdMain.Rows - 1, 5) = Val(rs.Fields("Sales_Tax") & "")
                grdMain.TextMatrix(grdMain.Rows - 1, 6) = Val(rs.Fields("Tax_Type") & "")
                grdMain.TextMatrix(grdMain.Rows - 1, 8) = Trim(rs.Fields("Extra") & "")
                grdMain.TextMatrix(grdMain.Rows - 1, 17) = Trim(rs.Fields("User_Overide") & "")
                If grdMain.ValueMatrix(grdMain.Rows - 1, 17) <> 0 Then
                    Reason = Round((grdMain.ValueMatrix(grdMain.Rows - 1, 17) - Int(grdMain.ValueMatrix(grdMain.Rows - 1, 17))) * 10, 0)
                    UserNo = Int(grdMain.ValueMatrix(grdMain.Rows - 1, 17))
                    ActiveReadServer1 "Select First_Name,Last_Name from Users where User_No = " & Val(UserNo)
                    If rs1.RecordCount > 0 Then
                        grdMain.TextMatrix(grdMain.Rows - 1, 17) = UserNo & " - " & rs1.Fields("First_Name") & " " & rs1.Fields("Last_Name")
                    End If
                    rs1.Close
                    ActiveReadServer1 "Select * from Reasons where Reason_No = " & Reason
                    If rs1.RecordCount > 0 Then
                        grdMain.TextMatrix(grdMain.Rows - 1, 17) = grdMain.TextMatrix(grdMain.Rows - 1, 17) & " (" & rs1.Fields("Reason_Name") & ")"
                    End If
                    rs1.Close
                Else
                    grdMain.TextMatrix(grdMain.Rows - 1, 17) = ""
                End If
                If rs.Fields("Function_Key") = 16 Then
                    If grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Service Charge" Then
                        SCharge = SCharge + rs.Fields("Line_Total")
                        TillData.Tipp = TillData.Tipp + rs.Fields("Line_Total")
                    End If
                End If
                If rs.Fields("Extra") & "" <> "" Then
                    Select Case rs.Fields("Extra") & ""
                        Case "Void"
                            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0C0FF
                        Case "Corr"
                            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0C0FF
                            grdMain.Cell(flexcpFontStrikethru, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = True
                        Case "Return Item"
                            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0C0FF
                        Case "Wastage"
                            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0C0FF
                        Case Else
                            If rs.Fields("Extra") & "" <> "" And rs.Fields("Function_Key") <> 20 Then
                                If Left(rs.Fields("Extra") & "", 6) = "Voided" Then
                                    grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = vbRed
                                Else
                                    grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = vbBlue
                                    grdMain.TextMatrix(grdMain.Rows - 1, 1) = "   >" & grdMain.TextMatrix(grdMain.Rows - 1, 1)
                                    grdMain.TextMatrix(grdMain.Rows - 1, 0) = ""
                                End If
                            Else
                                grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Extra") & " (Rate " & rs.Fields("Product_Code") & ")"
                            End If
                    End Select
                End If
                grdMain.TextMatrix(grdMain.Rows - 1, 9) = rs.Fields("Product_Code") & ""
                grdMain.TextMatrix(grdMain.Rows - 1, 10) = rs.Fields("Department_No") & ""
                lblTable.Caption = ""
                If rs.Fields("Table_No") <> 0 Then
                    lblTable.Caption = "Table No: " & rs.Fields("Table_No") & " Covers: " & rs.Fields("Covers")
                    TillData.TableNo = rs.Fields("Table_No")
                    TillData.Covers = rs.Fields("Covers")
                End If
                If rs.Fields("Tab_No") <> 0 Then
                    lblTable.Caption = "Bar Tab"
                End If
                
                TillData.DocNo = rs.Fields("Invoice_No")
                lblAcc.Caption = ""
                Select Case rs.Fields("Function_Key")
                    Case 9
                        grdMain.TextMatrix(grdMain.Rows - 1, 0) = ""
                        grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Cash Tendered"
                        grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0FFFF
                        If grdMain.TextMatrix(grdMain.Rows - 1, 2) = "" Then grdMain.TextMatrix(grdMain.Rows - 1, 2) = "0.00"
                        If SCharge <> 0 Then
                            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 2) + SCharge, "0.00")
                            SCharge = 0
                        End If
                        If rs.Fields("Dicount_Value") <> 0 Then grdMain.Rows = grdMain.Rows + 1: grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Discount Value": grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Dicount_Value"), "0.00")
                    Case 10
                        grdMain.TextMatrix(grdMain.Rows - 1, 0) = ""
                        grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Card Tendered"
                        grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0FFFF
                        If grdMain.TextMatrix(grdMain.Rows - 1, 2) = "" Then grdMain.TextMatrix(grdMain.Rows - 1, 2) = "0.00"
                        If SCharge <> 0 Then
                            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 2) + SCharge, "0.00")
                            SCharge = 0
                        End If
                    Case 11
                        grdMain.TextMatrix(grdMain.Rows - 1, 0) = ""
                        grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Voucher Tendered"
                        grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0FFFF
                        If grdMain.TextMatrix(grdMain.Rows - 1, 2) = "" Then grdMain.TextMatrix(grdMain.Rows - 1, 2) = "0.00"
                        If SCharge <> 0 Then
                            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 2) + SCharge, "0.00")
                            SCharge = 0
                        End If
                    Case 12
                        grdMain.TextMatrix(grdMain.Rows - 1, 0) = ""
                        grdMain.TextMatrix(grdMain.Rows - 1, 1) = "Charge Tendered"
                        grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HC0FFFF
                        ActiveReadServer2 "Select Debtor_Name from Debtors where Debtor_No = '" & rs.Fields("Account_No") & "'"
                        If rs2.RecordCount > 0 Then
                            lblAcc.Caption = "Charged to: " & rs.Fields("Account_No") & " - " & rs2.Fields("Debtor_Name")
                        Else
                            lblAcc.Caption = ""
                        End If
                        rs2.Close
                        If grdMain.TextMatrix(grdMain.Rows - 1, 2) = "" Then grdMain.TextMatrix(grdMain.Rows - 1, 2) = "0.00"
                        If SCharge <> 0 Then
                            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(grdMain.ValueMatrix(grdMain.Rows - 1, 2) + SCharge, "0.00")
                            SCharge = 0
                        End If
                    Case 16
                        grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 17) = &HEEFBE8
                        grdMain.TextMatrix(grdMain.Rows - 1, 0) = ""
                End Select
            End If
            rs.MoveNext
        Wend
    End If
    rs.Close
    grdMain.Row = grdMain.Rows - 1
    grdMain.ShowCell grdMain.Rows - 1, 0
    Sale_Total = 0
    Sale_Total_Excl = 0
    
    For i = 1 To grdMain.Rows - 1
        If grdMain.TextMatrix(i, 8) <> "Corr" Then
            If grdMain.ValueMatrix(i, 0) <> 0 Then
                Sale_Total = Val(Sale_Total) + Val(grdMain.TextMatrix(i, 2))
                Sale_Total_Excl = Val(Sale_Total_Excl) + Val(grdMain.TextMatrix(i, 2) / ((100 + grdMain.ValueMatrix(i, 5)) / 100))
            Else
                If grdMain.TextMatrix(i, 1) = "Service Charge" Then
                    grdMain.TextMatrix(i, 0) = ""
                    Sale_Total = Val(Sale_Total) + Val(grdMain.TextMatrix(i, 2))
                End If
            End If
        End If
    Next i
    lblSubTotal.Caption = Format(Sale_Total_Excl, "0.00") & " "
    lblTotal.Caption = Format(Sale_Total, "0.00") & " "
    lblVat.Caption = Format(Sale_Total - Sale_Total_Excl - TillData.Tipp, "0.00") & " "
    ActiveReadServer2 "Select User_Overide,(Select User_Name from Users where Transaction_View.User_Overide = Users.User_No) as User1 from Transaction_View where Function_Key = 14 and Invoice_No= " & Val(Table_No)
    If rs2.RecordCount > 0 Then
        txtowner.Caption = rs2.Fields("User_Overide") & " - " & rs2.Fields("User1")
    End If
    rs2.Close
    Screen.MousePointer = 0
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    grdMain.Rows = 1
    grdMain.TextMatrix(0, 0) = " No"
    grdMain.TextMatrix(0, 1) = "Description"
    grdMain.TextMatrix(0, 2) = "Total "
    grdMain.TextMatrix(0, 17) = "User Override"
    grdMain.ColWidth(0) = 400
    grdMain.ColWidth(1) = 2570
    grdMain.ColWidth(2) = 1060
    grdMain.ColWidth(14) = 100
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignRightCenter
    grdMain.ColHidden(3) = True
    grdMain.ColHidden(4) = True
    grdMain.ColHidden(5) = True
    grdMain.ColHidden(6) = True
    grdMain.ColHidden(7) = True
    grdMain.ColHidden(8) = True
    grdMain.ColHidden(9) = True
    grdMain.ColHidden(10) = True
    grdMain.ColHidden(11) = True
    grdMain.ColHidden(12) = True
    grdMain.ColHidden(13) = True
    grdMain.ColHidden(14) = True
    grdMain.ColHidden(15) = True
    grdMain.ColHidden(16) = True
    grdMain.ColHidden(17) = False
    grdMain.Cell(flexcpForeColor, 0, 0, 0, 17) = &H808080
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TillData.SaleTotal = 0
    TillData.TaxTotal = 0
    TillData.DocNo = 0
    TillData.Covers = 0
    TillData.Tipp = 0
End Sub

Private Sub grdMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If InStr(grdMain.TextMatrix(NewRow, 1), "Tendered") = 0 Then
        cmdTender.Enabled = False
    Else
        cmdTender.Enabled = True
    End If
End Sub

