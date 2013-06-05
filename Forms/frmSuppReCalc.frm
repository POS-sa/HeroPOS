VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSuppReCalc 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   510
      TabIndex        =   0
      Top             =   2700
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   767
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   810
      TabIndex        =   2
      Top             =   2430
      Width           =   5265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recalculate Supplier Balances"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   420
      TabIndex        =   1
      Top             =   1770
      Width           =   6135
   End
   Begin MSForms.Image Image1 
      Height          =   1725
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   7005
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   3
      Size            =   "12356;3043"
      Picture         =   "frmSuppReCalc.frx":0000
   End
   Begin MSForms.Image Image1 
      Height          =   1125
      Index           =   1
      Left            =   0
      Top             =   2190
      Width           =   7005
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;1984"
   End
   Begin MSForms.Image Image1 
      Height          =   555
      Index           =   2
      Left            =   0
      Top             =   1680
      Width           =   7005
      BackColor       =   15523287
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;979"
   End
End
Attribute VB_Name = "frmSuppReCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    DoEvents
    Recalc
End Sub
Private Sub Recalc()
    On Error Resume Next
    ProgressBar1.Value = 0
    ActiveReadServer "Select Supplier_No,Supplier_Name from Suppliers order by Supplier_No"
    ProgressBar1.Max = rs.RecordCount
    While Not rs.EOF
        ProgressBar1.Value = ProgressBar1.Value + 1
        DoEvents
        lblCap.Caption = "Recalculating " & rs.Fields("Supplier_No") & " - " & rs.Fields("Supplier_Name")
        ActiveReadServer2 "Select * from Supplier_Accounts where Account_No = '" & rs.Fields("Supplier_No") & "' order by Date_Time"
        Balance = 0
        While Not rs2.EOF
            Balance = Balance + rs2.Fields("Debit") - rs2.Fields("Credit")
            ActiveUpdateServer "Update Supplier_Accounts set Balance = " & Balance & " Where Line_no = " & rs2.Fields("Line_No")
            rs2.MoveNext
        Wend
        rs2.Close
        ActiveUpdateServer "Update Suppliers set Balance = " & Balance & " Where Supplier_no = '" & rs.Fields("Supplier_No") & "'"
        rs.MoveNext
    Wend
    rs.Close
    ProgressBar1.Value = ProgressBar1.Max
    MsgBox "Recalulation Completed Succesfully"
    Unload Me
    On Error Resume Next
End Sub

