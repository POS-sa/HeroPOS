VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmReCalc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recalculate Recipe Costs..."
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "frmReCalc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   510
      TabIndex        =   0
      Top             =   2670
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   714
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSForms.Image Image1 
      Height          =   1755
      Index           =   0
      Left            =   0
      Top             =   -30
      Width           =   7005
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   3
      Size            =   "12356;3096"
      Picture         =   "frmReCalc.frx":000C
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recipe Cost Updater"
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
      TabIndex        =   2
      Top             =   1770
      Width           =   6135
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   810
      TabIndex        =   1
      Top             =   2400
      Width           =   5265
   End
   Begin MSForms.Image Image1 
      Height          =   525
      Index           =   2
      Left            =   0
      Top             =   1680
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
      Top             =   2160
      Width           =   7005
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;1931"
   End
End
Attribute VB_Name = "frmReCalc"
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
    ActiveReadServer1 "Select Recipe_Item,Product_Code,Description,Unit_of_Measure,Unit_Size,Ave_Cost from Products where Stock_Item = 1 or Recipe_Item=1 and Unit_of_Measure <> 'Preparation Recipe' order by Recipe_Item"
    ProgressBar1.Max = rs1.RecordCount
    While Not rs1.EOF
        ProgressBar1.Value = ProgressBar1.Value + 1
        Label1.Caption = "Updating " & ProgressBar1.Value & " of " & ProgressBar1.Max
        ActiveReadServer "Select * from Recipes where Line_Code = '" & rs1.Fields("Product_Code") & "' order by Product_Code"
        resCost = 0
        If rs1.Fields("Recipe_Item") = 1 Then
            lblCap.Caption = "Updating " & rs1.Fields("Product_Code") & " - " & rs1.Fields("Description")
            ActiveReadServer2 "Select sum(Cost) as Ave_Cost from Recipes where Line_Type not in (6,7) and Product_code = '" & rs1.Fields("Product_Code") & "'"
            DoEvents
            If rs2.RecordCount > 0 Then
                ActiveUpdateServer "Update Products set Landed_Cost=" & rs2.Fields("Ave_Cost") & ",Ave_Cost = " & rs2.Fields("Ave_Cost") & " where Product_Code = '" & rs1.Fields("Product_Code") & "'"
            End If
            rs2.Close
        End If
        While Not rs.EOF
            lblCap.Caption = "Updating " & rs.RecordCount & " Recipes"
            Factor1 = 1
            Select Case UCase(rs1.Fields("Unit_of_Measure") & " to " & rs.Fields("Unit_of_Measure"))
                Case "ML TO SINGLE TOT"
                    Select Case rs1.Fields("Unit_Size")
                        Case "1000"
                            Factor1 = 1000 / 25
                            resCost = rs1.Fields("Ave_Cost") / Factor1
                        Case "750"
                            Factor1 = 750 / 25
                            resCost = rs1.Fields("Ave_Cost") / Factor1
                        Case "500"
                            Factor1 = 500 / 25
                            resCost = rs1.Fields("Ave_Cost") / Factor1
                    End Select
                Case "ML TO DOUBLE TOT"
                    Select Case rs1.Fields("Unit_Size")
                        Case "1000"
                            Factor1 = 1000 / 50
                            resCost = rs1.Fields("Ave_Cost") / Factor1
                        Case "750"
                            Factor1 = 750 / 50
                            resCost = rs1.Fields("Ave_Cost") / Factor1
                        Case "500"
                            Factor1 = 500 / 50
                            resCost = rs1.Fields("Ave_Cost") / Factor1
                    End Select
                Case "ML TO LT"
                Case "LT TO ML"
                    If Val(rs1.Fields("Unit_Size") & "") = 0 Then
                        Unit = 1
                    Else
                        Unit = rs1.Fields("Unit_Size")
                    End If
                    Factor1 = (Unit * 1000) / rs.Fields("Qty_Used")
                    resCost = rs1.Fields("Ave_Cost") / Factor1
                Case "KG TO G"
                    If Val(rs1.Fields("Unit_Size") & "") = 0 Then
                        Unit = 1
                    Else
                        Unit = rs1.Fields("Unit_Size")
                    End If
                    Factor1 = (Unit * 1000) / rs.Fields("Qty_Used")
                    resCost = rs1.Fields("Ave_Cost") / Factor1
                Case "G TO KG"
                Case "ML TO ML", "G TO G", "LT TO LT", "KG TO KG"
                    Factor1 = rs1.Fields("Unit_Size") / rs.Fields("Qty_Used")
                    resCost = rs1.Fields("Ave_Cost") / Factor1
                Case Else
                    resCost = rs1.Fields("Ave_Cost") * rs.Fields("Qty_Used")
            End Select
            ActiveUpdateServer "Update Recipes set Cost = " & resCost & " where Line_No = " & rs.Fields("Line_No")
            DoEvents
            rs.MoveNext
        Wend
        rs.Close
        DoEvents
        rs1.MoveNext
    Wend
    rs1.Close
    MsgBox "Recalulation Completed Succesfully"
    Unload Me
    On Error Resume Next
End Sub

