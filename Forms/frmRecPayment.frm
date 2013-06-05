VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRecPayment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Receive Payment..."
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmRecPayment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInvoice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   630
      Width           =   3735
   End
   Begin VB.TextBox txtGrvNo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   3120
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid grdPay 
      Height          =   1680
      Left            =   150
      TabIndex        =   0
      Top             =   1170
      Width           =   5295
      _cx             =   9340
      _cy             =   2963
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
      BackColorSel    =   16744576
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   8421504
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   1500
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRecPayment.frx":000C
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
      Editable        =   2
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
   Begin btButtonEx.ButtonEx cmdClose 
      Height          =   345
      Left            =   4380
      TabIndex        =   4
      Top             =   3090
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Don't Pay"
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
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   345
      Left            =   3150
      TabIndex        =   5
      Top             =   3090
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Appearance      =   3
      Caption         =   "Done"
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No: "
      Height          =   225
      Left            =   2610
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin MSForms.Image Image3 
      Height          =   345
      Left            =   3990
      Top             =   150
      Width           =   1455
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "2566;609"
   End
   Begin MSForms.Image Image12 
      Height          =   345
      Left            =   1530
      Top             =   540
      Width           =   3915
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "6906;609"
   End
   Begin MSForms.Image Image11 
      Height          =   345
      Left            =   1530
      Top             =   150
      Width           =   1455
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "2566;609"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Debtor Details: "
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   570
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No: "
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin MSForms.Image Image1 
      Height          =   1935
      Left            =   60
      Top             =   1050
      Width           =   5505
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "9710;3413"
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Total: "
      Height          =   225
      Left            =   0
      TabIndex        =   6
      Top             =   3180
      Width           =   1455
   End
   Begin MSForms.Image Image2 
      Height          =   345
      Left            =   1470
      Top             =   3090
      Width           =   1575
      BorderColor     =   8421504
      BackColor       =   16777215
      Size            =   "2778;609"
   End
   Begin MSForms.Image Image4 
      Height          =   915
      Left            =   60
      Top             =   60
      Width           =   5505
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "9710;1614"
   End
End
Attribute VB_Name = "frmRecPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonEx1_Click()
    ActiveReadServer1 "Select isnull(max(Payment_No),0)+1 as Receipt_No from Debtor_Accounts where Transaction_Type = 'Receipt'"
    Payment_No = rs1.Fields("Receipt_No")
    rs1.Close
    Payment = 0
    Balance = 0
    Tender_Type = ""
    Ref_No = ""
    ActiveReadServer "Select Balance from Debtor_Accounts where Account_No = '" & txtGrvNo & "' order by Line_No"
    If rs.RecordCount > 0 Then
        rs.MoveLast
        Balance = rs.Fields("Balance")
    End If
    rs.Close
    For i = 1 To grdPay.Rows - 1
        If grdPay.ValueMatrix(i, 1) <> 0 Then
            Payment = grdPay.ValueMatrix(i, 1)
            Select Case i
                Case 1: Tender_Type = "Cash"
                Case 2: Tender_Type = "Voucher"
                Case 3: Tender_Type = "Card"
                Case 4: Tender_Type = "EFT"
            End Select
            Ref_No = grdPay.TextMatrix(i, 2)
            Exit For
        End If
    Next i
    If Tender_Type = "" Then
        MsgBox "You have not entered a Payment Amount", vbCritical
        Unload Me
        Exit Sub
    End If
    If frmRecPayment.Tag = "Debtor" Then
        ActiveUpdateServer "INSERT INTO [Debtor_Accounts]([User_No],[Date_Time],[Transaction_Type], [Invoice_No], [Account_No], [Debit], [Credit], [Balance],[Tender_Type],[Ref_No],[Payment_No])" & _
        "VALUES(" & UserRecord.User_Number & ",Getdate(),'Receipt','" & txtInvoice.Text & "','" & txtGrvNo.Text & "',0," & Payment & "," & Balance - (Payment) & ",'" & Tender_Type & "','" & Ref_No & "'," & Payment_No & ")"
        DoEvents
        ActiveUpdateServer "Update Debtors set Balance= Balance - " & Payment & " where Debtor_No='" & txtGrvNo.Text & "'"
    End If
    Unload Me
End Sub
Private Sub Form_Activate()
    grdPay.TextMatrix(0, 0) = "Payment Type"
    grdPay.TextMatrix(0, 1) = "Value"
    grdPay.TextMatrix(0, 2) = "Ref No"
    grdPay.TextMatrix(1, 0) = "Cash Payout"
    grdPay.TextMatrix(2, 0) = "Voucher"
    grdPay.TextMatrix(3, 0) = "Card"
    grdPay.TextMatrix(4, 0) = "EFT"
    grdPay.TextMatrix(1, 1) = "0.00"
    grdPay.TextMatrix(2, 1) = "0.00"
    grdPay.TextMatrix(3, 1) = "0.00"
    grdPay.TextMatrix(4, 1) = "0.00"
    If frmRecPayment.Tag = "Debtor" Then
        txtGrvNo.Text = frmAccount.lblAcc.Caption
        txtDescription.Text = frmAccount.lblName.Caption
        txtTotal.Text = frmAccount.grdAcc.TextMatrix(frmAccount.grdAcc.Row, 2)
        Label1.Caption = "Account No: "
        Label3.Caption = "Invoice Total: "
        txtInvoice.Text = frmAccount.grdAcc.TextMatrix(frmAccount.grdAcc.Row, 5)
    End If
End Sub
Private Sub grdPay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If grdPay.Col = 1 Then
        If grdPay.ValueMatrix(grdPay.Row, grdPay.Col) <> Val(txtTotal.Text) Then
            Response = MsgBox("Are you sure you want to Pay this Amount?" & Chr(13) & "As it does not match the Invoice Total.", vbYesNo, "HeroPOS")
            If Response = vbNo Then Exit Sub
        End If
        grdPay.TextMatrix(grdPay.Row, grdPay.Col) = Format(grdPay.TextMatrix(grdPay.Row, grdPay.Col), "0.00")
    End If
    For i = 1 To grdPay.Rows - 1
        If i <> grdPay.Row Then
            grdPay.TextMatrix(i, 1) = "0.00"
        End If
    Next i
End Sub
Private Sub grdPay_KeyDown(KeyCode As Integer, Shift As Integer)
With grdPay
    Select Case KeyCode
        Case 13 'Enter
        Case 46
            Select Case grdPay.Col
                Case 1
                    grdPay.TextMatrix(grdPay.Row, grdPay.Col) = "0.00"
                Case 2
                    grdPay.TextMatrix(grdPay.Row, grdPay.Col) = ""
            End Select
        Case 45, 48 To 57, 96 To 105, 109, 110, 189
            Select Case grdPay.Col
                Case 1, 2
                    .EditCell
            End Select
        Case Else
            If grdPay.Col = 2 Then
                .EditCell
            End If
    End Select
End With
End Sub
Private Sub grdPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case 1
            If InStr(grdPay.EditSelText, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
            Select Case KeyAscii
                Case 8, 13, 27, 46, 48 To 57
                Case Else: KeyAscii = 0
            End Select
        Case 2
            Select Case KeyAscii
                Case 39
                    KeyAscii = Asc("`")
            End Select
    End Select
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

