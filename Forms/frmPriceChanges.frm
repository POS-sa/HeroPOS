VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPriceChanges 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price Changes"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   Icon            =   "frmPriceChanges.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUnitSize 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   21
      Top             =   8080
      Width           =   2295
   End
   Begin VB.TextBox txtUnitoM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   20
      Top             =   7740
      Width           =   2295
   End
   Begin VB.TextBox txtShort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   19
      Top             =   7380
      Width           =   3855
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   18
      Top             =   7040
      Width           =   6645
   End
   Begin VB.TextBox txtProductCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   17
      Top             =   6690
      Width           =   2355
   End
   Begin VB.TextBox txtSellIncl 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1695
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   8835
      Width           =   1935
   End
   Begin VB.TextBox txtLandCost 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1695
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   8490
      Width           =   765
   End
   Begin VB.TextBox txtCCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2865
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   8490
      Width           =   915
   End
   Begin VSFlex8Ctl.VSFlexGrid grdGrid 
      Height          =   6480
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8245
      _cx             =   14543
      _cy             =   11430
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColorSel    =   16706532
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPriceChanges.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin btButtonEx.ButtonEx cmdFin 
      Height          =   540
      Left            =   5490
      TabIndex        =   13
      Top             =   8550
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   953
      Appearance      =   3
      Caption         =   " Finalize Selected Prices"
      CaptionOffsetY  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdPrintCurrent 
      Height          =   570
      Left            =   5490
      TabIndex        =   14
      Top             =   7950
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   1005
      Appearance      =   3
      Caption         =   "Print labels for Current Product Only"
      CaptionOffsetY  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   1140
      Left            =   3960
      TabIndex        =   2
      Text            =   "1"
      Top             =   7950
      Width           =   1485
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   5850
      Top             =   6690
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   3990
      Top             =   6690
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unknown Product"
      Height          =   255
      Left            =   6210
      TabIndex        =   16
      Top             =   6750
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Product"
      Height          =   255
      Left            =   4350
      TabIndex        =   15
      Top             =   6750
      Width           =   1965
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   4
      Left            =   1590
      Top             =   8775
      Width           =   2295
      BackColor       =   16777215
      Size            =   "4048;529"
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   3
      Left            =   1590
      Top             =   8430
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;529"
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   8100
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Unit Size:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCost 
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   8460
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
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   8820
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Selling Price:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   6750
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
      Index           =   10
      Left            =   120
      TabIndex        =   8
      Top             =   7785
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
      Index           =   11
      Left            =   120
      TabIndex        =   7
      Top             =   7440
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
      Index           =   12
      Left            =   120
      TabIndex        =   6
      Top             =   7095
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "Description:"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   225
      Left            =   4050
      TabIndex        =   5
      Top             =   7710
      Width           =   1395
      BackColor       =   -2147483643
      Caption         =   "No Off Labels"
      Size            =   "2461;397"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Image Image5 
      Height          =   300
      Index           =   0
      Left            =   2760
      Top             =   8430
      Width           =   1125
      BackColor       =   16777215
      Size            =   "1984;529"
   End
   Begin MSForms.Image Image1 
      Height          =   2595
      Left            =   60
      Top             =   6600
      Width           =   8265
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "14579;4577"
   End
End
Attribute VB_Name = "frmPriceChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DrvName As String
Private Sub cmdFin_Click()
    For i = 1 To grdGrid.Rows - 1
        If grdGrid.TextMatrix(i, 3) <> "" Then
            ActiveReadServer1 "Select * from Products_Replicate where Product_Code = '" & grdGrid.TextMatrix(i, 0) & "'"
            If rs1.RecordCount > 0 Then
                ActiveReadServer2 "Select * from Products where Product_Code = '" & rs1.Fields("Product_Code") & "'"
                If rs2.RecordCount > 0 Then
                    ActiveUpdateServer "Update Products set " & _
                    " [Description]='" & rs1.Fields("Description") & "'," & _
                    " [Short_Description]='" & rs1.Fields("Short_Description") & "'," & _
                    " [Department_No]='" & rs1.Fields("Department_No") & "'," & _
                    " [Pack_Size]='" & rs1.Fields("Pack_Size") & "'," & _
                    " [Unit_Size]='" & rs1.Fields("Unit_Size") & "'," & _
                    " [Unit_of_Measure]='" & rs1.Fields("Unit_of_Measure") & "'," & _
                    " [Maximum_Discount]='" & rs1.Fields("Maximum_Discount") & "'," & _
                    " [Sales_Item]='" & rs1.Fields("Sales_Item") & "'," & _
                    " [Stock_Item]='" & rs1.Fields("Stock_Item") & "'," & _
                    " [Returnable_Item]='" & rs1.Fields("Returnable_Item") & "'," & _
                    " [Recipe_Item]='" & rs1.Fields("Recipe_Item") & "'," & _
                    " [Touch_Item]='" & rs1.Fields("Touch_Item") & "'," & _
                    " [Scale_Item]='" & rs1.Fields("Scale_Item") & "'," & _
                    " [Whole_Unit]='" & rs1.Fields("Whole_Unit") & "'," & _
                    " [Sales_Tax]='" & rs1.Fields("Sales_Tax") & "'," & _
                    " [Tax_Type]='" & rs1.Fields("Tax_Type") & "'," & _
                    " [Once_off]='" & rs1.Fields("Once_off") & "'," & _
                    " [Date_Updated]=Getdate() where Product_Code = '" & rs1.Fields("Product_Code") & "'"
                    DoEvents
                    ActiveUpdateServer "Delete from Products_Replicate where Product_Code = '" & rs1.Fields("Product_Code") & "'"
                Else
                    ActiveUpdateServer "Insert into Products ([Product_Code], [Description], [Short_Description], [Department_No], [Pack_Size], [Unit_Size], [Unit_of_Measure], [Maximum_Discount], [Sales_Item], [Stock_Item], [Returnable_Item], [Recipe_Item], [Touch_Item], [Scale_Item], [Whole_Unit], [Sales_Tax], [Tax_Type], [Once_off], [Date_Created], [Date_Updated]) " & _
                    " values ('" & rs1.Fields("Product_Code") & "','" & rs1.Fields("Description") & "','" & rs1.Fields("Short_Description") & "','" & rs1.Fields("Department_No") & "','" & rs1.Fields("Pack_Size") & "','" & rs1.Fields("Unit_Size") & "','" & rs1.Fields("Unit_of_Measure") & "','" & rs1.Fields("Maximum_Discount") & "','" & rs1.Fields("Sales_Item") & "','" & rs1.Fields("Stock_Item") & "','" & rs1.Fields("Returnable_Item") & "','" & rs1.Fields("Recipe_Item") & "','" & rs1.Fields("Touch_Item") & "','" & rs1.Fields("Scale_Item") & "','" & rs1.Fields("Whole_Unit") & "','" & rs1.Fields("Sales_Tax") & "','" & rs1.Fields("Tax_Type") & "','" & rs1.Fields("Once_Off") & "',Getdate(),Getdate()" & ")"
                    DoEvents
                    ActiveUpdateServer "Delete from Products_Replicate where Product_Code = '" & rs1.Fields("Product_Code") & "'"
                End If
                rs2.Close
            End If
            rs1.Close
            DoEvents
            ActiveUpdateServer "Update Products set Selling_Price = " & grdGrid.TextMatrix(i, 2) & ",Landed_Cost  = " & grdGrid.TextMatrix(i, 4) & ",Ave_Cost  = " & grdGrid.TextMatrix(i, 4) & " where Product_Code = '" & grdGrid.TextMatrix(i, 0) & "'"
            DoEvents
            ActiveUpdateServer "Delete from Price_Changes where Product_Code = '" & grdGrid.TextMatrix(i, 0) & "'"
            DoEvents
        End If
    Next i
    grdGrid.Rows = 1
    ActiveReadServer2 "Select Cost_Price,Product_Code,Retail_Price," & _
    "  isnull((Select Description from Products where Products.Product_Code = Price_Changes.Product_Code)," & _
    " isnull('N|' +(Select Description from Products_Replicate where Products_Replicate.Product_Code = Price_Changes.Product_Code group by Description)," & _
    " 'Unknown Product')) as Description from Price_Changes order by Line_No"
    While Not rs2.EOF
        grdGrid.Rows = grdGrid.Rows + 1
        grdGrid.TextMatrix(grdGrid.Rows - 1, 0) = rs2.Fields("Product_Code")
        If Left(rs2.Fields("Description") & "", 2) = "N|" Then
            grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = Mid(rs2.Fields("Description") & "", 3)
            grdGrid.Cell(flexcpFontBold, grdGrid.Rows - 1, 0, grdGrid.Rows - 1, 4) = True
            grdGrid.Cell(flexcpForeColor, grdGrid.Rows - 1, 0, grdGrid.Rows - 1, 4) = &HFF0000
        Else
            If rs2.Fields("Description") & "" = "Unknown Product" Then
                grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = "Unknown Product"
                grdGrid.Cell(flexcpFontBold, grdGrid.Rows - 1, 0, grdGrid.Rows - 1, 4) = True
                grdGrid.Cell(flexcpForeColor, grdGrid.Rows - 1, 0, grdGrid.Rows - 1, 4) = &HC0&
            Else
                grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = rs2.Fields("Description") & ""
            End If
        End If
        grdGrid.TextMatrix(grdGrid.Rows - 1, 2) = Format(rs2.Fields("Retail_Price"), "0.00")
        grdGrid.TextMatrix(grdGrid.Rows - 1, 4) = Format(rs2.Fields("Cost_Price"), "0.00")
        rs2.MoveNext
    Wend
    rs2.Close
    If grdGrid.Rows > 1 Then grdGrid.Row = 1
End Sub

Private Sub cmdPrintCurrent_Click()
      
    
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

Private Sub Form_Activate()
    If grdGrid.Rows > 1 Then
        grdGrid.Row = 1
        grdGrid.Col = 0
    End If
End Sub
Private Sub Form_Load()
    For Each prnPrinter In Printers
        If InStr(prnPrinter.DeviceName, Devices.Label_Printer) Then
            DrvName = prnPrinter.DeviceName
        End If
    Next
    grdGrid.ColHidden(4) = True
    grdGrid.TextMatrix(0, 0) = "Product Code"
    grdGrid.TextMatrix(0, 1) = "Description"
    grdGrid.TextMatrix(0, 2) = "Selling Price"
    grdGrid.TextMatrix(0, 3) = "Selected"
    grdGrid.ColWidth(0) = grdGrid.Width * 0.2
    grdGrid.ColWidth(1) = grdGrid.Width * 0.5
    grdGrid.ColWidth(2) = grdGrid.Width * 0.15
    grdGrid.ColWidth(3) = grdGrid.Width * 0.15
    grdGrid.ColAlignment(0) = flexAlignLeftCenter
    grdGrid.ColAlignment(1) = flexAlignLeftCenter
    grdGrid.ColAlignment(2) = flexAlignRightCenter
    grdGrid.ColAlignment(3) = flexAlignCentreCenter
    grdGrid.ColDataType(3) = flexDTBoolean
    grdGrid.Rows = 1
    ActiveReadServer2 "Select Cost_Price,Product_Code,Retail_Price," & _
    "  isnull((Select Description from Products where Products.Product_Code = Price_Changes.Product_Code)," & _
    " isnull('N|' +(Select Description from Products_Replicate where Products_Replicate.Product_Code = Price_Changes.Product_Code group by Description)," & _
    " 'Unknown Product')) as Description from Price_Changes order by Line_No"
    While Not rs2.EOF
        grdGrid.Rows = grdGrid.Rows + 1
        grdGrid.TextMatrix(grdGrid.Rows - 1, 0) = rs2.Fields("Product_Code")
        If Left(rs2.Fields("Description") & "", 2) = "N|" Then
            grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = Mid(rs2.Fields("Description") & "", 3)
            grdGrid.Cell(flexcpFontBold, grdGrid.Rows - 1, 0, grdGrid.Rows - 1, 4) = True
            grdGrid.Cell(flexcpForeColor, grdGrid.Rows - 1, 0, grdGrid.Rows - 1, 4) = &HFF0000
        Else
            If rs2.Fields("Description") & "" = "Unknown Product" Then
                grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = "Unknown Product"
                grdGrid.Cell(flexcpFontBold, grdGrid.Rows - 1, 0, grdGrid.Rows - 1, 4) = True
                grdGrid.Cell(flexcpForeColor, grdGrid.Rows - 1, 0, grdGrid.Rows - 1, 4) = &HC0&
            Else
                grdGrid.TextMatrix(grdGrid.Rows - 1, 1) = rs2.Fields("Description") & ""
            End If
        End If
        grdGrid.TextMatrix(grdGrid.Rows - 1, 2) = Format(rs2.Fields("Retail_Price"), "0.00")
        grdGrid.TextMatrix(grdGrid.Rows - 1, 4) = Format(rs2.Fields("Cost_Price"), "0.00")
        rs2.MoveNext
    Wend
    rs2.Close
    If grdGrid.Rows > 1 Then grdGrid.Row = 1
End Sub
Private Sub grdGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    ActiveReadServer "Select * from Products where Product_code ='" & grdGrid.TextMatrix(NewRow, 0) & "'"
    If rs.RecordCount > 0 Then
        txtProductCode.Text = rs.Fields("Product_Code")
        txtDescription.Text = rs.Fields("Description")
        txtShort.Text = Format(rs.Fields("Short_Description"), "0.00")
        txtUnitoM.Text = rs.Fields("Unit_of_Measure")
        If Val(rs.Fields("Unit_Size") & "") = 0 Then
            txtUnitSize.Text = ""
        Else
            txtUnitSize.Text = rs.Fields("Unit_Size")
        End If
        txtLandCost.Text = Format(grdGrid.TextMatrix(NewRow, 4), "0.00")
        txtSellIncl.Text = Format(grdGrid.TextMatrix(NewRow, 2), "0.00")
    Else
        txtProductCode.Text = grdGrid.TextMatrix(NewRow, 0)
        txtDescription.Text = grdGrid.TextMatrix(NewRow, 1)
        txtShort.Text = grdGrid.TextMatrix(NewRow, 1)
        txtUnitoM.Text = "each"
        txtUnitSize.Text = ""
        txtLandCost.Text = Format(grdGrid.TextMatrix(NewRow, 4), "0.00")
        txtSellIncl.Text = Format(grdGrid.TextMatrix(NewRow, 2), "0.00")
    End If
End Sub
Private Sub grdGrid_EnterCell()
    Select Case grdGrid.Col
        Case 0, 1, 2
            grdGrid.Editable = flexEDNone
        Case 3
            If grdGrid.TextMatrix(grdGrid.Row, 1) = "Unknown Product" Then
                grdGrid.Editable = flexEDNone
            Else
                grdGrid.Editable = flexEDKbdMouse
            End If
    End Select
End Sub
Private Sub cmdPrint_Click()
    Dim byBuf(50000) As Byte
    Dim tempStr As String
    Dim Count As Integer
    tempStr = ""
    Count = 0
    tempStr = tempStr + "CB" + Chr(13)
    tempStr = tempStr + "SS3" + Chr(13)
    tempStr = tempStr + "SD20" + Chr(13)
    tempStr = tempStr + "SOT" + Chr(13)
    tempStr = tempStr + "SW832" + Chr(13)
    tempStr = tempStr + "SL" & Devices.Barcode_Height & ",20,G" + Chr(13)
    tempStr = tempStr + "T194,6,2,0,0,0,0,N,N,'" & txtDescription.Text & "'" + Chr(13)
    tempStr = tempStr + "T445,110,3,0,0,0,0,N,N,'R " & txtSellIncl.Text & "'" + Chr(13)
    tempStr = tempStr + "T445,70,1,0,0,0,0,N,N,'" & txtCCode.Text & "'" + Chr(13)
    tempStr = tempStr + "B1224,38,0,2,4,80,0,1,'" & txtProductCode.Text & "'" + Chr(13)
    tempStr = tempStr + "P" & Text1.Text & "" + Chr(13)

    'Converting string to byte type
    Count = Count + UniStringToByte(tempStr, byBuf, Count)
        
    nRet = DirectWrite(DrvName, byBuf(0), Count)
    If nRet <> SEM_SUCCESS Then
        ErrorMessage (nRet)
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

