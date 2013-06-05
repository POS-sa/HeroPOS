VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmClear 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmClear.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   570
      TabIndex        =   1
      Top             =   2970
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   767
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   870
      TabIndex        =   2
      Top             =   2700
      Width           =   5265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "System Data Clear"
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
      TabIndex        =   0
      Top             =   1950
      Width           =   6135
   End
   Begin MSForms.Image Image1 
      Height          =   555
      Index           =   2
      Left            =   60
      Top             =   1860
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
      Top             =   2460
      Width           =   7005
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;1984"
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
      Picture         =   "frmClear.frx":000C
   End
End
Attribute VB_Name = "frmClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    ProgressBar1.Max = 18
    ProgressBar1.Value = 0
    
    lblCap.Caption = "Clearing Sales Journal"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Sales_Journal"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Consumption Journal"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Consumption_Journal"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Counters"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Counters"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Tabs"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Tab_Listing"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Tables"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Table_Listing"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Quantities"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Quantities"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Stock Takes"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Stock_Take_Journal"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Stock Take Listing"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Stock_Take_Listing"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing User Journal"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from User_Journal"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Transfer Journal"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Transfer_Journal"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Purchase Journal"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Purchase_Journal"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Supplier Accounts"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Supplier_Accounts"
    DoEvents
    DoEvents
    ActiveUpdateServer "Update Suppliers set Balance=0"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Saved GRV's"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Purchase_Journal_Listing"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Room Accounts"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Room_Accounts"
    DoEvents
    ActiveUpdateServer "Update Debtors set Balance=0"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing Debtor Accounts"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Debtor_Accounts"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing all Reservations"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Reservations"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clearing all Print Journals"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Delete from Print_journal"
    stime = Timer
    While Timer - stime < 1: Wend
    
    lblCap.Caption = "Clocking all Users Out"
    ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    ActiveUpdateServer "Update Users Set Drawer_No = 0,Clocked_In= 0"
    stime = Timer
    While Timer - stime < 1: Wend
    ProgressBar1.Value = 18
    DoEvents
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No,Navigate) values (" & UserRecord.User_Number & ",Getdate(),15," & Workstation_No & ",'System Wide Totals Clear')"
    MsgBox "Data Clear Completed. Please Restart the Application", vbInformation, "HeroPOS Message"
    Unload Me
End Sub

