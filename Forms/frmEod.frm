VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmEod 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4410
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin btButtonEx.ButtonEx cmdRun 
      Height          =   465
      Left            =   1320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3810
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   820
      Appearance      =   3
      Caption         =   "Run End-of-Day"
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
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   570
      TabIndex        =   1
      Top             =   3030
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   767
      _Version        =   327682
      Appearance      =   1
   End
   Begin btButtonEx.ButtonEx cmdEnd 
      Height          =   465
      Left            =   3540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3810
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   820
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
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   870
      TabIndex        =   3
      Top             =   2760
      Width           =   5265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "End of Day Run"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1980
      Width           =   6135
   End
   Begin MSForms.Image Image1 
      Height          =   1755
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   7005
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   3
      Size            =   "12356;3096"
      Picture         =   "frmEod.frx":0000
   End
   Begin MSForms.Image Image1 
      Height          =   555
      Index           =   2
      Left            =   60
      Top             =   1890
      Width           =   7005
      BackColor       =   15523287
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;979"
   End
   Begin MSForms.Image Image1 
      Height          =   1125
      Index           =   1
      Left            =   60
      Top             =   2520
      Width           =   7005
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;1984"
   End
End
Attribute VB_Name = "frmEod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
    Unload Me
End Sub
Private Sub cmdRun_Click()
    Run_Eod
    Unload Me
End Sub
Private Sub Run_Eod()
    ActiveReadServer1 "Select isnull(max(EOD_No),0)+1 as Eod_No from EOD"
    EOD_No = rs1.Fields("EOD_No")
    rs1.Close
    ActiveReadServer "Select *,(Select Ave_Cost from Products where Quantities.Product_Code = Products.Product_Code) as Ave_Cost from Quantities"
    ProgressBar1.Max = rs.RecordCount
    While Not rs.EOF
        lblCap.Caption = "Processing Record no: " & ProgressBar1.Value & " of " & ProgressBar1.Max
        DoEvents
        ActiveUpdateServer "Insert into EOD (Product_Code, Location_No, Stock_on_Hand,Date_Time,EOD_No,User_No,Ave_Cost) " & _
        "values ('" & rs.Fields("product_Code") & "','" & rs.Fields("Location_No") & "','" & rs.Fields("Stock_on_Hand") & "',Getdate()," & EOD_No & "," & UserRecord.User_Number & "," & rs.Fields("Ave_Cost") & ")"
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Wend
    ProgressBar1.Value = rs.RecordCount - 1
    rs.Close
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),5," & Workstation_No & ",'EOD Run')"
    MsgBox "End-of-Day Run Completed.", vbInformation, "HeroPOS"
End Sub

