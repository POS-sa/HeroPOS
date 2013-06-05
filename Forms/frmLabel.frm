VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLabel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Label Printing"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7365
   Icon            =   "frmLabel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3990
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   210
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtCCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2805
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1890
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   3930
      TabIndex        =   0
      Text            =   "1"
      Top             =   1290
      Width           =   1845
   End
   Begin VB.TextBox txtLandCost 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1635
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   1890
      Width           =   765
   End
   Begin VB.TextBox txtSellIncl 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1635
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   2265
      Width           =   1935
   End
   Begin btButtonEx.ButtonEx cmdPrint 
      Height          =   585
      Left            =   5820
      TabIndex        =   1
      Top             =   1290
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1032
      Appearance      =   3
      Caption         =   "Print Price Labels"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdPrint1 
      Height          =   585
      Left            =   5820
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1032
      Appearance      =   3
      Caption         =   "Print Shelf Talkers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   0
      Left            =   2700
      Top             =   1830
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;529"
   End
   Begin MSForms.TextBox txtUnitoM 
      Height          =   285
      Left            =   1530
      TabIndex        =   16
      Top             =   1140
      Width           =   2295
      VariousPropertyBits=   746604567
      MaxLength       =   25
      BorderStyle     =   1
      Size            =   "4048;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   225
      Left            =   4020
      TabIndex        =   15
      Top             =   1020
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "No of Labels"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   12
      Left            =   60
      TabIndex        =   14
      Top             =   495
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Description:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   11
      Left            =   60
      TabIndex        =   13
      Top             =   840
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Button Description:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   10
      Left            =   60
      TabIndex        =   12
      Top             =   1185
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Unit of Measure:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   6
      Left            =   60
      TabIndex        =   11
      Top             =   150
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Product Code:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   5
      Left            =   60
      TabIndex        =   10
      Top             =   2220
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Selling Price:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCost 
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   1860
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Landed Cost:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   1500
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Unit Size:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtDescription 
      Height          =   285
      Left            =   1530
      TabIndex        =   7
      Top             =   435
      Width           =   5745
      VariousPropertyBits=   746604567
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "10134;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtShort 
      Height          =   285
      Left            =   1530
      TabIndex        =   6
      Top             =   780
      Width           =   2295
      VariousPropertyBits=   746604567
      MaxLength       =   25
      BorderStyle     =   1
      Size            =   "4048;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtUnitSize 
      Height          =   285
      Left            =   1530
      TabIndex        =   5
      Top             =   1485
      Width           =   2295
      VariousPropertyBits=   746604567
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "4048;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtProductCode 
      Height          =   285
      Left            =   1530
      TabIndex        =   4
      Top             =   90
      Width           =   2325
      VariousPropertyBits=   748701719
      MaxLength       =   16
      BorderStyle     =   1
      Size            =   "4101;503"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   3
      Left            =   1530
      Top             =   1830
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;529"
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   4
      Left            =   1530
      Top             =   2205
      Width           =   2295
      BackColor       =   16777215
      Size            =   "4048;529"
   End
End
Attribute VB_Name = "frmLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DrvName As String
Private Sub cmdPrint_Click()
    Dim byBuf(50000) As Byte
    Dim tempStr As String
    Dim Count As Integer
    Dim wwww As Integer
    tempStr = ""
    Count = 0
wwww = (Devices.Label_Width / 5)
'    tempStr = tempStr + "CB" + Chr(13)
'    tempStr = tempStr + "SS3" + Chr(13)
'    tempStr = tempStr + "SD20" + Chr(13)
'    tempStr = tempStr + "SOT" + Chr(13)
'    tempStr = tempStr + "SW832" + Chr(13)
'    tempStr = tempStr + "SL" & Devices.Barcode_Height & ",20,G" + Chr(13)
'    tempStr = tempStr + "T194,6,2,0,0,0,0,N,N,'" & txtDescription.Text & "'" + Chr(13)
'    tempStr = tempStr + "T445,110,3,0,0,0,0,N,N,'R " & txtSellIncl.Text & "'" + Chr(13)
'    tempStr = tempStr + "T445,70,1,0,0,0,0,N,N,'" & txtCCode.Text & "'" + Chr(13)
'    tempStr = tempStr + "B1224,38,0,2,4,80,0,1,'" & txtProductCode.Text & "'" + Chr(13)
'    tempStr = tempStr + "P" & Text1.Text & "" + Chr(13)


    tempStr = tempStr + "CB" + Chr(13)
    tempStr = tempStr + "SS3" + Chr(13)
    tempStr = tempStr + "SD20" + Chr(13)
    tempStr = tempStr + "SOT" + Chr(13)
    tempStr = tempStr + "SW400" + Chr(13)
    tempStr = tempStr + "SL" & Devices.Barcode_Height & ",20,G" + Chr(13)
    tempStr = tempStr + "T80,6,2,0,0,0,0,N,N,'" & txtDescription.Text & "'" + Chr(13)
    tempStr = tempStr + "T80,110,3,0,0,0,0,N,N,' " & txtSellIncl.Text & "'" + Chr(13)
    tempStr = tempStr + "T80,70,1,0,0,0,0,N,N,'" & txtCCode.Text & "'" + Chr(13)
    tempStr = tempStr + "B1" & wwww & ",38,1,2,3,50,0,1,'" & txtProductCode.Text & "'" + Chr(13)
    tempStr = tempStr + "P" & Text1.Text & "" + Chr(13)




    'Converting string to byte type
    Count = Count + UniStringToByte(tempStr, byBuf, Count)
        
    nRet = DirectWrite(DrvName, byBuf(0), Count)
    If nRet <> SEM_SUCCESS Then
        ErrorMessage (nRet)
    End If

End Sub

Private Sub cmdPrint1_Click()
    Dim byBuf(50000) As Byte
    Dim tempStr As String
    Dim Count As Integer
    tempStr = ""
    Count = 0
    Devices.Label_Height = 252
    tempStr = tempStr + "CB" + Chr(13)
    tempStr = tempStr + "SS3" + Chr(13)
    tempStr = tempStr + "SD20" + Chr(13)
    tempStr = tempStr + "SOT" + Chr(13)
    tempStr = tempStr + "SW832" + Chr(13)
    'tempStr = tempStr + "SL" & Devices.Barcode_Height & ",20,G" + Chr(13)
    tempStr = tempStr + "SL" & Devices.Label_Height & ",20,G" + Chr(13)
    If Val(txtPack.Text) > 1 Then
        tempStr = tempStr + "T194,0,2,0,0,0,0,N,N,'" & txtDescription.Text & " (Pack of " & txtPack.Text & ")'" + Chr(13)
    Else
        tempStr = tempStr + "T194,0,2,0,0,0,0,N,N,'" & txtDescription.Text & "'" + Chr(13)
    End If
    tempStr = tempStr + "T594,160,2,0,0,0,0,N,N,'VAT INCL'" + Chr(13)
    tempStr = tempStr + "T195,110,6,1,1,1,0,N,N,'R " & txtSellIncl.Text & "'" + Chr(13)
    tempStr = tempStr + "T445,70,1,0,0,0,0,N,N,'" & txtCCode.Text & "'" + Chr(13)
    tempStr = tempStr + "B1194,28,0,2,4,50,0,1,'" & txtProductCode.Text & "'" + Chr(13)
    tempStr = tempStr + "P" & Text1.Text & "" + Chr(13)

    'Converting string to byte type
    Count = Count + UniStringToByte(tempStr, byBuf, Count)
        
    nRet = DirectWrite(DrvName, byBuf(0), Count)
    If nRet <> SEM_SUCCESS Then
        ErrorMessage (nRet)
    End If
End Sub
Private Sub Form_Activate()
    cmdPrint.Enabled = False
    
    Select Case frmLabel.Tag
        Case "GRV"
            With frmGRV
                ActiveReadServer2 "Select * from Products where Product_Code ='" & .grdGRV.TextMatrix(.grdGRV.Row, 0) & "'"
                If rs2.RecordCount > 0 Then
                    txtProductCode.Text = rs2.Fields("Product_Code")
                    txtDescription.Text = rs2.Fields("Description")
                    txtShort.Text = rs2.Fields("Short_Description")
                    txtUnitoM.Text = rs2.Fields("Unit_of_Measure")
                    txtUnitSize.Text = rs2.Fields("Unit_Size")
                    txtSellIncl.Text = Format(rs2.Fields("Selling_Price"), "0.00")
                    txtPack.Text = rs2.Fields("Pack_Size")
                End If
                rs2.Close
                Text1.Text = .grdGRV.TextMatrix(.grdGRV.Row, 5)
                txtLandCost.Text = Format(.grdGRV.TextMatrix(.grdGRV.Row, 7), "0.00")
                cmdPrint.Enabled = True
                cmdPrint1.Enabled = True
             End With
        Case "PriceChange"
            With frmPriceChange
                ActiveReadServer2 "Select * from Products where Product_Code ='" & .txtProductCode & "'"
                If rs2.RecordCount > 0 Then
                    txtProductCode.Text = rs2.Fields("Product_Code")
                    txtDescription.Text = rs2.Fields("Description")
                    txtShort.Text = rs2.Fields("Short_Description")
                    txtUnitoM.Text = rs2.Fields("Unit_of_Measure")
                    txtUnitSize.Text = rs2.Fields("Unit_Size")
                    txtSellIncl.Text = Format(rs2.Fields("Selling_Price"), "0.00")
                    txtLandCost.Text = Format(rs2.Fields("Landed_Cost"), "0.00")
                    txtPack.Text = rs2.Fields("Pack_Size")
                End If
                rs2.Close
                Text1.Text = "1"
                cmdPrint.Enabled = True
                cmdPrint1.Enabled = True
             End With
        Case "PriceChangeRow"
            With frmPriceChange
                ActiveReadServer2 "Select * from Products where Product_Code ='" & .grdGrid.TextMatrix(.grdGrid.Row, 0) & "'"
                If rs2.RecordCount > 0 Then
                    txtProductCode.Text = rs2.Fields("Product_Code")
                    txtDescription.Text = rs2.Fields("Description")
                    txtShort.Text = rs2.Fields("Short_Description")
                    txtUnitoM.Text = rs2.Fields("Unit_of_Measure")
                    txtUnitSize.Text = rs2.Fields("Unit_Size")
                    txtSellIncl.Text = .txtSellIncl.Text
                    txtLandCost.Text = Format(rs2.Fields("Landed_Cost"), "0.00")
                    txtPack.Text = rs2.Fields("Pack_Size")
                End If
                rs2.Close
                Text1.Text = "1"
                cmdPrint.Enabled = True
                cmdPrint1.Enabled = True
             End With
        Case Else
            With frmProducts
                txtProductCode.Text = .txtProductCode.Text
                txtDescription.Text = .txtDescription.Text
                txtShort.Text = .txtShort.Text
                txtUnitoM.Text = .cmbUnit.Text
                txtUnitSize.Text = .txtUnitSize.Text
                txtLandCost.Text = Format(.txtLandCost.Text, "0.00")
                txtSellIncl.Text = Format(.txtSellIncl.Text, "0.00")
                txtPack.Text = .txtPackSize.Text
             End With
             ActiveReadServer2 "Select sum(Stock_on_Hand) as Stock_on_Hand from Quantities where Product_Code = '" & txtProductCode.Text & "'"
             If rs2.RecordCount > 0 Then
                If Val(rs2.Fields("Stock_on_Hand") & "") > 0 Then
                    cmdPrint.Enabled = True
                     cmdPrint1.Enabled = True
                Else
                    cmdPrint.Enabled = False
                    
                End If
             End If
             rs2.Close
     End Select
     If txtUnitSize.Text = "0" Then txtUnitSize.Text = ""
End Sub

Private Sub Form_Load()
    
    cmdPrint.Enabled = False
    
    For Each prnPrinter In Printers
        If InStr(prnPrinter.DeviceName, Devices.Label_Printer) Then
            DrvName = prnPrinter.DeviceName
        End If
    Next
    If DrvName = "" Then
        MsgBox ("SRP-770 Driver not installed.")
        Unload Me
        Exit Sub
    End If
End Sub

Sub ErrorMessage(ByVal errcode As Long)
    If errcode = SEM_ERR_NOPRINTER Then
        msgtext = "Specified printer driver does not exist"
    ElseIf errcode = SEM_ERR_NOTSUPPORT Then
        msgtext = "Specified printer or port are not supported"
    ElseIf errcode = SEM_ERR_OPEN Then
        msgtext = "Cannot open printer port"
    ElseIf errcode = SEM_ERR_WRITE Then
        msgtext = "Write Error"
    ElseIf errcode = SEM_ERR_READ Then
        msgtext = "Read Error"
    ElseIf errcode = SEM_ERR_TIMEOUT Then
        msgtext = "Timeout Error"
    ElseIf errcode = SEM_ERR_PARAM Then
        msgtext = "Function Parameter Error"
    End If
    
    Call MsgBox(msgtext, vbCritical, "API ERROR")
End Sub
Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtLandCost_Change()
    txtCCode.Text = ""
    For i = 1 To Len(txtLandCost.Text)
        Select Case Mid(txtLandCost.Text, i, 1)
            Case "1": txtCCode.Text = txtCCode.Text & Cost_Code.One
            Case "2": txtCCode.Text = txtCCode.Text & Cost_Code.Two
            Case "3": txtCCode.Text = txtCCode.Text & Cost_Code.Three
            Case "4": txtCCode.Text = txtCCode.Text & Cost_Code.Four
            Case "5": txtCCode.Text = txtCCode.Text & Cost_Code.Five
            Case "6": txtCCode.Text = txtCCode.Text & Cost_Code.Six
            Case "7": txtCCode.Text = txtCCode.Text & Cost_Code.Seven
            Case "8": txtCCode.Text = txtCCode.Text & Cost_Code.Eight
            Case "9": txtCCode.Text = txtCCode.Text & Cost_Code.Nine
            Case "0": txtCCode.Text = txtCCode.Text & Cost_Code.Ten
            Case ".": txtCCode.Text = txtCCode.Text & "."
        End Select
    Next i
End Sub

