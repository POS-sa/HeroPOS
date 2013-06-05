VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTender 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " New Tender Type...."
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1890
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   1890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdCash 
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1296
      Appearance      =   3
      BackColor       =   12648447
      Caption         =   "Cash"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdCard 
      Height          =   735
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1296
      Appearance      =   3
      BackColor       =   12648447
      Caption         =   "Card"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdCharge 
      Height          =   735
      Left            =   60
      TabIndex        =   3
      Top             =   1620
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1296
      Appearance      =   3
      BackColor       =   12648447
      Caption         =   "Charge"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdCheque 
      Height          =   735
      Left            =   60
      TabIndex        =   4
      Top             =   2400
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1296
      Appearance      =   3
      BackColor       =   12648447
      Caption         =   "Voucher"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx ButtonEx5 
      Height          =   735
      Left            =   60
      TabIndex        =   5
      Top             =   3180
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1296
      Appearance      =   3
      BackColor       =   8438015
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VB.PictureBox picHold 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   30
      ScaleHeight     =   585
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   510
      Width           =   285
   End
End
Attribute VB_Name = "frmTender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx5_Click()
    Unload Me
End Sub
Private Sub cmdCard_Click()
    Select Case frmTransView.grdMain.TextMatrix(frmTransView.grdMain.Row, 1)
        Case "Cash Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 9 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 10 where function_Key = 9 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Cash_Sales_Value = isnull(Cash_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) - 1" & _
                    ",Card_Sales_Value = isnull(Card_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Cash_Sales_Value = isnull(Cash_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) - 1" & _
                    ",Card_Sales_Value = isnull(Card_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
            End If
            rs.Close
        Case "Voucher Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 11 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 10 where function_Key = 11 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) - 1" & _
                    ",Card_Sales_Value = isnull(Card_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) - 1" & _
                    ",Card_Sales_Value = isnull(Card_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
            End If
            rs.Close
        Case "Charge Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 12 and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 10,Account_No = '' where function_Key = 12 and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Charge_Sales_Value = isnull(Charge_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) - 1" & _
                    ",Card_Sales_Value = isnull(Card_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Charge_Sales_Value = isnull(Charge_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) - 1" & _
                    ",Card_Sales_Value = isnull(Card_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
                ActiveUpdateServer "Delete from Debtor_Accounts where Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                    
                ActiveReadServer2 "Select * from Debtor_Accounts where Account_No = '" & rs.Fields("Account_No") & "' order by Date_Time"
                Balance = 0
                While Not rs2.EOF
                    Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
                    ActiveUpdateServer "Update Debtor_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
                    rs2.MoveNext
                Wend
                rs2.Close
                ActiveUpdateServer "Update Debtors set Balance = " & Balance & " Where Debtor_no = '" & rs.Fields("Account_No") & "'"
            End If
            rs.Close
    End Select
    MsgBox "Tender Type Changed to Card." & Chr(13) & "Please note that the users Cashup Totals will be Changed as Well!", vbInformation, "HeroPOS"
    Unload Me
End Sub
Private Sub cmdCash_Click()
    Select Case frmTransView.grdMain.TextMatrix(frmTransView.grdMain.Row, 1)
        Case "Card Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 10 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 9 where function_Key = 10 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Card_Sales_Value = isnull(Card_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) - 1" & _
                    ",Cash_Sales_Value = isnull(Cash_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Card_Sales_Value = isnull(Card_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) - 1" & _
                    ",Cash_Sales_Value = isnull(Cash_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
            End If
            rs.Close
        Case "Voucher Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 11 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 9 where function_Key = 11 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) - 1" & _
                    ",Cash_Sales_Value = isnull(Cash_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) - 1" & _
                    ",Cash_Sales_Value = isnull(Cash_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
            End If
            rs.Close
        Case "Charge Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 12 and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 9,Account_No = '' where function_Key = 12 and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Charge_Sales_Value = isnull(Charge_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) - 1" & _
                    ",Cash_Sales_Value = isnull(Cash_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Charge_Sales_Value = isnull(Charge_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) - 1" & _
                    ",Cash_Sales_Value = isnull(Cash_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
                ActiveUpdateServer "Delete from Debtor_Accounts where Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                    
                ActiveReadServer2 "Select * from Debtor_Accounts where Account_No = '" & rs.Fields("Account_No") & "' order by Date_Time"
                Balance = 0
                While Not rs2.EOF
                    Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
                    ActiveUpdateServer "Update Debtor_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
                    rs2.MoveNext
                Wend
                rs2.Close
                ActiveUpdateServer "Update Debtors set Balance = " & Balance & " Where Debtor_no = '" & rs.Fields("Account_No") & "'"
            End If
            rs.Close
    End Select
    MsgBox "Tender Type Changed to Cash." & Chr(13) & "Please note that the users Cashup Totals will be Changed as Well!", vbInformation, "HeroPOS"
    Unload Me
End Sub
Private Sub cmdCharge_Click()
    frmDebTrans.Show vbModal
    Select Case frmDebTrans.Tag
        Case ""
            MsgBox "Tender Type not Changed.", vbInformation, "HeroPOS"
            Unload Me
            Exit Sub
        Case Else
            Debtor_No = frmDebTrans.Tag
    End Select
    Unload frmDebTrans
    Select Case frmTransView.grdMain.TextMatrix(frmTransView.grdMain.Row, 1)
        Case "Card Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 10 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 12 where function_Key = 10 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Card_Sales_Value = isnull(Card_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) - 1" & _
                    ",Charge_Sales_Value = isnull(Charge_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Card_Sales_Value = isnull(Card_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) - 1" & _
                    ",Charge_Sales_Value = isnull(Charge_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
                SaleTot = rs.Fields("Line_Total")
            End If
            rs.Close
        Case "Cash Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 9 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key =12 where function_Key = 9 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Cash_Sales_Value = isnull(Cash_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) - 1" & _
                    ",Charge_Sales_Value = isnull(Charge_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Cash_Sales_Value = isnull(Cash_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) - 1" & _
                    ",Charge_Sales_Value = isnull(Charge_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
                SaleTot = rs.Fields("Line_Total")
            End If
            rs.Close
        Case "Voucher Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 11 and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 12,Account_No = '' where function_Key = 11 and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) - 1" & _
                    ",Charge_Sales_Value = isnull(Charge_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) - 1" & _
                    ",Charge_Sales_Value = isnull(Charge_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
                DoEvents
                SaleTot = rs.Fields("Line_Total")
            End If
            rs.Close
    End Select
    ActiveUpdateServer "Update Debtors set Balance=Balance + " & SaleTot & " where Debtor_No = '" & Trim(Debtor_No) & "'"
    ActiveReadServer "Select isnull(Balance,0) as Balance from Debtors where Debtor_No = '" & Trim(Debtor_No) & "'"
    If rs.RecordCount > 0 Then
        NewBalance = rs.Fields("Balance")
    Else
        NewBalance = 0
    End If
    rs.Close
    ActiveUpdateServer "INSERT INTO [Debtor_Accounts]([Transaction_Type],[Date_Time], [Invoice_No], [Account_No], [Debit], [Credit], [Balance])" & _
    "VALUES('Invoice',Getdate()," & Val(frmTransView.lblDocNo) & ",'" & Trim(Debtor_No) & "'," & SaleTot & ",0," & NewBalance & ")"
    MsgBox "Tender Type Changed to Charge." & Chr(13) & "Please note that the users Cashup Totals will be Changed as Well!", vbInformation, "HeroPOS"
    Unload Me
End Sub
Private Sub cmdCheque_Click()
    Select Case frmTransView.grdMain.TextMatrix(frmTransView.grdMain.Row, 1)
        Case "Card Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 10 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 11 where function_Key = 10 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Card_Sales_Value = isnull(Card_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) - 1" & _
                    ",Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Card_Sales_Value = isnull(Card_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Card_Sales_Qty = isnull(Card_Sales_Qty,0) - 1" & _
                    ",Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
            End If
            rs.Close
        Case "Cash Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 9 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key =11 where function_Key = 9 and Line_Total = " & frmTransView.grdMain.ValueMatrix(frmTransView.grdMain.Row, 2) & " and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Cash_Sales_Value = isnull(Cash_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) - 1" & _
                    ",Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Cash_Sales_Value = isnull(Cash_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Cash_Sales_Qty = isnull(Cash_Sales_Qty,0) - 1" & _
                    ",Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
            End If
            rs.Close
        Case "Charge Tendered"
            ActiveReadServer "Select * from Sales_Journal where function_Key = 12 and Invoice_No = " & Val(frmTransView.lblDocNo)
            If rs.RecordCount > 0 Then
                ActiveUpdateServer "Update Sales_Journal set Function_Key = 11,Account_No = '' where function_Key = 12 and Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                ActiveReadServer1 "Select * from Counters where User_No = " & rs.Fields("User_No") & " and (Open_Trans_No - 1 < " & Val(frmTransView.lblDocNo) & " AND Close_Trans_No + 1 > " & Val(frmTransView.lblDocNo) & ")"
                If rs1.RecordCount = 1 Then
                    ActiveUpdateServer " Update Counters set" & _
                    " Charge_Sales_Value = isnull(Charge_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) - 1" & _
                    ",Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) + 1" & _
                    " where Cashup_No = " & rs1.Fields("Cashup_No")
                Else
                    ActiveUpdateServer " Update Counters set" & _
                    " Charge_Sales_Value = isnull(Charge_Sales_Value,0) - " & rs.Fields("Line_Total") & _
                    ",Charge_Sales_Qty = isnull(Charge_Sales_Qty,0) - 1" & _
                    ",Cheque_Sales_Value = isnull(Cheque_Sales_Value,0) + " & rs.Fields("Line_Total") & _
                    ",Cheque_Sales_Qty = isnull(Cheque_Sales_Qty,0) + 1" & _
                    " where Close_Trans_No is null and User_No = " & rs.Fields("User_No")
                End If
                rs1.Close
                ActiveUpdateServer "Delete from Debtor_Accounts where Invoice_No = " & Val(frmTransView.lblDocNo)
                DoEvents
                    
                ActiveReadServer2 "Select * from Debtor_Accounts where Account_No = '" & rs.Fields("Account_No") & "' order by Date_Time"
                Balance = 0
                While Not rs2.EOF
                    Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
                    ActiveUpdateServer "Update Debtor_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
                    rs2.MoveNext
                Wend
                rs2.Close
                ActiveUpdateServer "Update Debtors set Balance = " & Balance & " Where Debtor_no = '" & rs.Fields("Account_No") & "'"
            End If
            rs.Close
    End Select
    MsgBox "Tender Type Changed to Voucher." & Chr(13) & "Please note that the users Cashup Totals will be Changed as Well!", vbInformation, "HeroPOS"
    Unload Me
End Sub
Private Sub Form_Load()
    If frmTransView.grdMain.TextMatrix(frmTransView.grdMain.Row, 1) = "Cash Tendered" Then
        cmdCash.Enabled = False
    End If
    If frmTransView.grdMain.TextMatrix(frmTransView.grdMain.Row, 1) = "Card Tendered" Then
        cmdCard.Enabled = False
    End If
    If frmTransView.grdMain.TextMatrix(frmTransView.grdMain.Row, 1) = "Charge Tendered" Then
        cmdCharge.Enabled = False
    End If
    If frmTransView.grdMain.TextMatrix(frmTransView.grdMain.Row, 1) = "Voucher Tendered" Then
        cmdCheque.Enabled = False
    End If
End Sub
