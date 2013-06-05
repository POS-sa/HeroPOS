VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmReplicate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Master Replication...."
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   510
      TabIndex        =   0
      Top             =   2700
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   714
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   810
      TabIndex        =   2
      Top             =   2430
      Width           =   5265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master Replication"
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
      Height          =   345
      Left            =   420
      TabIndex        =   1
      Top             =   1800
      Width           =   6135
   End
   Begin MSForms.Image Image1 
      Height          =   1755
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   7005
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   3
      Size            =   "12356;3096"
      Picture         =   "frmReplicate.frx":0000
   End
   Begin MSForms.Image Image1 
      Height          =   525
      Index           =   2
      Left            =   0
      Top             =   1710
      Width           =   7005
      BackColor       =   15523287
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;926"
   End
   Begin MSForms.Image Image1 
      Height          =   1095
      Index           =   1
      Left            =   0
      Top             =   2190
      Width           =   7005
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;1931"
   End
End
Attribute VB_Name = "frmReplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    DoEvents
    Replicate
End Sub
Private Sub Replicate()
    On Error Resume Next
    ProgressBar1.Value = 0
    ActiveReadServer3 "Select * from Products_Replicate "
    ProgressBar1.Max = rs3.RecordCount
    While Not rs3.EOF
        ProgressBar1.Value = ProgressBar1.Value + 1
        lblCap.Caption = "Updating " & ProgressBar1.Value & " of " & ProgressBar1.Max
        ActiveReadServer2 "Select * from Products where Product_Code = '" & rs3.Fields("Product_Code") & "'"
        If rs2.RecordCount > 0 Then
            ActiveUpdateServer "Update Products set " & _
            " [Description]='" & rs3.Fields("Description") & "'," & _
            " [Short_Description]='" & rs3.Fields("Short_Description") & "'," & _
            " [Department_No]='" & rs3.Fields("Department_No") & "'," & _
            " [Pack_Size]='" & rs3.Fields("Pack_Size") & "'," & _
            " [Unit_Size]='" & rs3.Fields("Unit_Size") & "'," & _
            " [Unit_of_Measure]='" & rs3.Fields("Unit_of_Measure") & "'," & _
            " [Maximum_Discount]='" & rs3.Fields("Maximum_Discount") & "'," & _
            " [Sales_Item]='" & rs3.Fields("Sales_Item") & "'," & _
            " [Stock_Item]='" & rs3.Fields("Stock_Item") & "'," & _
            " [Returnable_Item]='" & rs3.Fields("Returnable_Item") & "'," & _
            " [Recipe_Item]='" & rs3.Fields("Recipe_Item") & "'," & _
            " [Touch_Item]='" & rs3.Fields("Touch_Item") & "'," & _
            " [Scale_Item]='" & rs3.Fields("Scale_Item") & "'," & _
            " [Whole_Unit]='" & rs3.Fields("Whole_Unit") & "'," & _
            " [Sales_Tax]='" & rs3.Fields("Sales_Tax") & "'," & _
            " [Tax_Type]='" & rs3.Fields("Tax_Type") & "'," & _
            " [Once_off]='" & rs3.Fields("Once_off") & "'," & _
            " [Date_Updated]=Getdate() where Product_Code = '" & rs3.Fields("Product_Code") & "'"
            DoEvents
            ActiveUpdateServer "Delete from Products_Replicate where Product_Code = '" & rs3.Fields("Product_Code") & "'"
        Else
            ActiveUpdateServer "Insert into Products ([Product_Code], [Description], [Short_Description], [Department_No], [Pack_Size], [Unit_Size], [Unit_of_Measure], [Maximum_Discount], [Sales_Item], [Stock_Item], [Returnable_Item], [Recipe_Item], [Touch_Item], [Scale_Item], [Whole_Unit], [Sales_Tax], [Tax_Type], [Once_off], [Date_Created], [Date_Updated]) " & _
            " values ('" & rs3.Fields("Product_Code") & "','" & rs3.Fields("Description") & "','" & rs3.Fields("Short_Description") & "','" & rs3.Fields("Department_No") & "','" & rs3.Fields("Pack_Size") & "','" & rs3.Fields("Unit_Size") & "','" & rs3.Fields("Unit_of_Measure") & "','" & rs3.Fields("Maximum_Discount") & "','" & rs3.Fields("Sales_Item") & "','" & rs3.Fields("Stock_Item") & "','" & rs3.Fields("Returnable_Item") & "','" & rs3.Fields("Recipe_Item") & "','" & rs3.Fields("Touch_Item") & "','" & rs3.Fields("Scale_Item") & "','" & rs3.Fields("Whole_Unit") & "','" & rs3.Fields("Sales_Tax") & "','" & rs3.Fields("Tax_Type") & "','" & rs3.Fields("Once_Off") & "',Getdate(),Getdate()" & ")"
            DoEvents
            ActiveUpdateServer "Delete from Products_Replicate where Product_Code = '" & rs3.Fields("Product_Code") & "'"
        End If
        rs2.Close
        DoEvents
        rs3.MoveNext
    Wend
    DoEvents
    rs3.Close
    Unload Me
    On Error Resume Next
End Sub

