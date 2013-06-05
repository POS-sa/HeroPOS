VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRecalcCon 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3255
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   510
      TabIndex        =   0
      Top             =   2670
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
      Top             =   2400
      Width           =   5265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recalculate Sale Consumption"
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
      Top             =   1740
      Width           =   6135
   End
   Begin MSForms.Image Image1 
      Height          =   1725
      Index           =   0
      Left            =   0
      Top             =   -30
      Width           =   7005
      BorderStyle     =   0
      SizeMode        =   1
      SpecialEffect   =   3
      Size            =   "12356;3043"
      Picture         =   "frmRecalcCon.frx":0000
   End
   Begin MSForms.Image Image1 
      Height          =   1125
      Index           =   1
      Left            =   0
      Top             =   2160
      Width           =   7005
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;1984"
   End
   Begin MSForms.Image Image1 
      Height          =   555
      Index           =   2
      Left            =   0
      Top             =   1650
      Width           =   7005
      BackColor       =   15523287
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "12356;979"
   End
End
Attribute VB_Name = "frmRecalcCon"
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
    ActiveUpdateServer "Delete from Consumption_Journal"
    ProgressBar1.Value = 0
    ActiveReadServer "Select * from Sales_Journal where Function_Key=7 and Product_Code <> '0' order by Line_No"
    ProgressBar1.Max = rs.RecordCount
    While Not rs.EOF
        ProgressBar1.Value = ProgressBar1.Value + 1
        lblCap.Caption = "Updating " & ProgressBar1.Value & " of " & ProgressBar1.Max
        Product_Code = rs.Fields("Product_Code")
        Qty = rs.Fields("Qty")
        Select Case Kitchen_Printer_No
            Case 0
                ActiveReadServer2 "Select Kitchen1 from Products where Product_Code = '" & Product_Code & "'"
                Kitchen_Printer = rs2.Fields("Kitchen1")
                rs2.Close
            Case 1
                ActiveReadServer2 "Select Kitchen2 from Products where Product_Code = '" & Product_Code & "'"
                Kitchen_Printer = rs2.Fields("Kitchen2")
                rs2.Close
        End Select
        Invoice_No = rs.Fields("Invoice_No")
        Ave_Cost = rs.Fields("Ave_Cost")
        NewLocation = rs.Fields("Location")
        If Product_Code <> "" Then
            'If Product_Code = "2112" Then Stop
            ActiveReadServer2 "Select Recipe_Item from Products where Recipe_Item = 1 and Product_Code = '" & Product_Code & "'"
            Select Case rs2.RecordCount
                Case 0 'No Recipe
                    Select Case Trim(Kitchen_Printer & "")
                        Case "<None>", ""
                        Case Else 'Location from Kitchen Printer
                            ActiveReadServer1 "Select * from Printer_Links where Printer = '" & Trim(Kitchen_Printer) & "'"
                            If rs1.RecordCount > 0 Then
                                If Val(rs1.Fields("Location_No") & "") <> 0 Then
                                    NewLocation = Val(rs1.Fields("Location_No") & "")
                                Else
                                    NewLocation = Location_No
                                End If
                            End If
                            rs1.Close
                    End Select
                    
                    
                    ActiveUpdateServer "Insert into Consumption_Journal (Product_Code,Location_No,Ave_Cost,Qty_Consumed,Date_Time,Invoice_No) values ('" & Product_Code & "'," & NewLocation & "," & Ave_Cost * Qty & "," & Qty & ",'" & rs.Fields("Date_Time") & "'," & Invoice_No & ")"
                Case 1 'Using a Recipe
                    Select Case Kitchen_Printer
                        Case "<None>", ""
                            SaveQty = Qty
                            ActiveReadServer1 "Select Line_Type,Unit_of_Measure as Recipe_Unit, Line_Code,Qty_Used,(Select Unit_Size from Products where Products.Product_code = Recipes.Line_Code) as Unit_Size, (Select Unit_of_Measure from Products where Products.Product_code = Recipes.Line_Code) as Unit_of_Measure,(Select Ave_Cost from Products where Products.Product_code = Recipes.Line_Code) as Ave_Cost from Recipes where Product_Code = '" & Product_Code & "' and Line_Type in (3,4)"
                            While Not rs1.EOF
                                Qty = SaveQty
                                Select Case rs1.Fields("Qty_Used")
                                    Case "1 x 25ml"
                                        Qty = Round(25 / rs1.Fields("Unit_Size"), 4) * Qty
                                    Case "2 x 25ml"
                                        Qty = Round(50 / rs1.Fields("Unit_Size"), 4) * Qty
                                    Case Else
                                        UnitSize = rs1.Fields("Unit_Size")
                                        If rs1.Fields("Unit_of_Measure") <> rs1.Fields("Recipe_Unit") Then
                                            Select Case UCase(rs1.Fields("Unit_of_Measure") & " to " & rs1.Fields("Recipe_Unit"))
                                                Case "ML TO LT"
                                                    UnitSize = rs1.Fields("Unit_Size") / 1000
                                                Case "LT TO ML"
                                                    If rs1.Fields("Unit_Size") = 0 Then
                                                        UnitSize = 1000
                                                    Else
                                                        UnitSize = rs1.Fields("Unit_Size") * 1000
                                                    End If
                                                Case "G TO KG"
                                                    UnitSize = rs1.Fields("Unit_Size") / 1000
                                                Case "KG TO G"
                                                    If rs1.Fields("Unit_Size") = 0 Then
                                                        UnitSize = 1000
                                                    Else
                                                        UnitSize = rs1.Fields("Unit_Size") * 1000
                                                    End If
                                                Case Else
                                                    UnitSize = rs1.Fields("Unit_Size")
                                            End Select
                                        End If
                                        If Val(UnitSize & "") <> 0 Then
                                            Qty = Round(rs1.Fields("Qty_Used") / UnitSize, 4) * Qty
                                        Else
                                            Qty = rs1.Fields("Qty_Used") * Qty
                                        End If
                                End Select
                                Ave_Cost = rs1.Fields("Ave_Cost") * Qty
                                ActiveUpdateServer "Insert into Consumption_Journal (Product_Code,Location_No,Ave_Cost,Qty_Consumed,Date_Time,Invoice_No) values ('" & rs1.Fields("Line_Code") & "'," & NewLocation & "," & Ave_Cost & "," & Qty & ",'" & rs.Fields("Date_Time") & "'," & Invoice_No & ")"
                                rs1.MoveNext
                            Wend
                            rs1.Close
                        Case Else 'Location from Kitchen Printer
                            ActiveReadServer1 "Select * from Printer_Links where Printer = '" & Trim(Kitchen_Printer) & "'"
                            If rs1.RecordCount > 0 Then
                                If Val(rs1.Fields("Location_No") & "") <> 0 Then
                                    NewLocation = Val(rs1.Fields("Location_No") & "")
                                Else
                                    NewLocation = Location_No
                                End If
                            End If
                            rs1.Close
                            ActiveReadServer1 "Select Unit_of_Measure as Recipe_Unit, Line_Code,Qty_Used,isnull((Select Unit_Size from Products where Products.Product_code = Recipes.Line_Code),0) as Unit_Size, (Select Unit_of_Measure from Products where Products.Product_code = Recipes.Line_Code) as Unit_of_Measure,(Select Ave_Cost from Products where Products.Product_code = Recipes.Line_Code) as Ave_Cost from Recipes where Product_Code = '" & Product_Code & "' and Line_Type in (3,4)"
                            SaveQty = Qty
                            While Not rs1.EOF
                                Qty = SaveQty
                                Select Case rs1.Fields("Qty_Used")
                                    Case "1 x 25ml"
                                        Qty = Round(25 / rs1.Fields("Unit_Size"), 4) * Qty
                                    Case "2 x 25ml"
                                        Qty = Round(50 / rs1.Fields("Unit_Size"), 4) * Qty
                                    Case Else
                                        UnitSize = rs1.Fields("Unit_Size")
                                        If rs1.Fields("Unit_of_Measure") <> rs1.Fields("Recipe_Unit") Then
                                            Select Case UCase(rs1.Fields("Unit_of_Measure") & " to " & rs1.Fields("Recipe_Unit"))
                                                Case "ML TO LT"
                                                    UnitSize = rs1.Fields("Unit_Size") / 1000
                                                Case "LT TO ML"
                                                    If rs1.Fields("Unit_Size") = 0 Then
                                                        UnitSize = 1000
                                                    Else
                                                        UnitSize = rs1.Fields("Unit_Size") * 1000
                                                    End If
                                                Case "G TO KG"
                                                    UnitSize = rs1.Fields("Unit_Size") / 1000
                                                Case "KG TO G"
                                                    If rs1.Fields("Unit_Size") = 0 Then
                                                        UnitSize = 1000
                                                    Else
                                                        UnitSize = rs1.Fields("Unit_Size") * 1000
                                                    End If
                                                Case Else
                                                    UnitSize = rs1.Fields("Unit_Size")
                                            End Select
                                        End If
                                        If Val(UnitSize & "") <> 0 Then
                                            Qty = Round(rs1.Fields("Qty_Used") / UnitSize, 4) * Qty
                                        Else
                                            Qty = rs1.Fields("Qty_Used") * Qty
                                        End If
                                End Select
                                Ave_Cost = rs1.Fields("Ave_Cost") * Qty
                                ActiveUpdateServer "Insert into Consumption_Journal (Product_Code,Location_No,Ave_Cost,Qty_Consumed,Date_Time,Invoice_No) values ('" & rs1.Fields("Line_Code") & "'," & NewLocation & "," & Ave_Cost & "," & Qty & ",'" & rs.Fields("Date_Time") & "'," & Invoice_No & ")"
                                rs1.MoveNext
                            Wend
                            rs1.Close
                    End Select
            End Select
            rs2.Close
        End If
        DoEvents
        rs.MoveNext
    Wend
    rs.Close
    MsgBox "Recalulation Completed Succesfully"
    Unload Me
    On Error Resume Next
End Sub

