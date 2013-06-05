VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmBackorders 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backorders"
   ClientHeight    =   9375
   ClientLeft      =   3750
   ClientTop       =   1995
   ClientWidth     =   14190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14190
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtBonumber 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   750
      Width           =   1605
   End
   Begin VSFlex8Ctl.VSFlexGrid grdBackorders 
      Height          =   6060
      Left            =   60
      TabIndex        =   0
      Top             =   2340
      Width           =   13560
      _cx             =   23918
      _cy             =   10689
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      BackColorSel    =   16642749
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16645618
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBackorders.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      ExplorerBar     =   5
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   1
         Top             =   11700
         Width           =   1005
      End
   End
   Begin btButtonEx.ButtonEx cmdSupplier 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   " Click to Search.... "
      Top             =   240
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Supplier..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   375
      Left            =   8220
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   8421504
      Caption         =   "¦"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin VB.PictureBox picBlocDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8700
      ScaleHeight     =   345
      ScaleWidth      =   2775
      TabIndex        =   5
      Top             =   240
      Width           =   2805
      Begin MSForms.Label lblDate 
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   90
         Width           =   2235
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "1 Feb 2009 to 2 Feb 2009"
         Size            =   "3942;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin RichTextLib.RichTextBox txtAddress 
      Height          =   765
      Left            =   4350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   750
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   1349
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmBackorders.frx":0078
   End
   Begin MSComCtl2.DTPicker mthViewStart 
      Height          =   375
      Left            =   8340
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy (ddd)"
      Format          =   65798147
      CurrentDate     =   39414
   End
   Begin btButtonEx.ButtonEx cmdOk 
      Height          =   330
      Left            =   10290
      TabIndex        =   13
      Top             =   1110
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Appearance      =   3
      Caption         =   "Ok"
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
   Begin MSComCtl2.DTPicker mthViewEnd 
      Height          =   375
      Left            =   10860
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy (ddd)"
      Format          =   65798147
      CurrentDate     =   39414
   End
   Begin btButtonEx.ButtonEx ButtonEx2 
      Height          =   285
      Left            =   8220
      TabIndex        =   7
      ToolTipText     =   " Click to Search.... "
      Top             =   690
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Appearance      =   3
      Caption         =   "Product..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   150
      TabIndex        =   15
      Top             =   720
      Width           =   435
   End
   Begin MSForms.Image Image2 
      Height          =   345
      Left            =   540
      Top             =   690
      Width           =   1785
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "3149;609"
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   4
      Left            =   60
      Top             =   8460
      Width           =   13575
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "23945;661"
   End
   Begin MSForms.Image Image6 
      Height          =   825
      Left            =   8220
      Top             =   660
      Visible         =   0   'False
      Width           =   5235
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "9234;1455"
   End
   Begin MSForms.Image Image10 
      Height          =   885
      Left            =   4230
      Top             =   660
      Width           =   3735
      BorderColor     =   12632256
      BackColor       =   16777215
      Size            =   "6588;1561"
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   165
      Left            =   3150
      TabIndex        =   11
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   2235
   End
   Begin MSForms.ComboBox Comboproducts 
      Height          =   285
      Left            =   9510
      TabIndex        =   8
      Tag             =   "Up"
      Top             =   660
      Width           =   3735
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6588;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      BorderColor     =   12632256
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbSuppliers 
      Height          =   285
      Left            =   4230
      TabIndex        =   4
      Tag             =   "Up"
      Top             =   240
      Width           =   3735
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6588;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      BorderColor     =   12632256
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   60
      Top             =   1860
      Width           =   13575
      BackColor       =   16051176
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "23945;661"
   End
   Begin MSForms.Image Image1 
      Height          =   1755
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   13545
      BackColor       =   16777215
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "23892;3096"
   End
End
Attribute VB_Name = "frmBackorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub grdBackorders_Click()
If grdBackorders.Tag = "newbackorder" Then
    
    If grdBackorders.TextMatrix(0, 0) = "User Number" And _
    grdBackorders.TextMatrix(0, 4) = "Product Code" Then
    grdBackorders.SelectionMode = flexSelectionFree
    grdBackorders.Editable = flexEDKbdMouse
            
            If grdBackorders.Rows = 1 Then
            grdBackorders.Rows = grdBackorders.Rows + 1
            grdBackorders.SetFocus
            End If
    End If
End If
End Sub
Public Sub SaveBackOrder()
Select Case frmBackorders.Tag
Case "Saveable"


Dim i As Integer
x = x + 1
ActiveUpdateServer " delete  * from backorder_journal where order_no = ' " & _
grdBackorders.TextMatrix(1, 3) & "'"

            For i = 1 To (grdBackorders.Rows - 1)
            
               locno = grdBackorders.TextMatrix(i, 1)
                
           
           
           ActiveReadServer2 " Select * from Products where Product_Code = '" & grdBackorders.TextMatrix(i, 4) & "'"
           'Donext
           If rs2.RecordCount > 0 Then
           Vat_Rate = rs2.Fields("Sales_Tax")
           Department_No = rs2.Fields("Department_No")
           Packsizes = rs2.Fields("Pack_Size")
           End If
           ActiveUpdateServer " Insert into Backorder_Journal (workstation_no , user_no, location_no, " & _
"Supplier_No, Order_No, Product_Code,Department_no, Pack_Size, Qty_Ordered, " & _
"Price_Ordered, Vat_Rate, Line_Total, Date_Time, Order_Date, Delivery_Date) " & _
" Values (" & Workstation_No & ", " & UserRecord.User_Number & _
", " & locno & ", '" & grdBackorders.TextMatrix(i, 2) & "', " & grdBackorders.TextMatrix(i, 3) & _
", " & grdBackorders.TextMatrix(i, 4) & ", '" & Department_No & "', " & Packsizes & ", " & grdBackorders.TextMatrix(i, 6) & _
", " & grdBackorders.TextMatrix(i, 7) & ", " & Vat_Rate & ", " & grdBackorders.TextMatrix(i, 7) & _
", Getdate(), '" & grdBackorders.TextMatrix(i, 8) & "', '" & grdBackorders.TextMatrix(i, 9) & "' ) "
            
            
            Next i

' finished so clear the grid and the supplier name and the order no
MsgBox "Backorder " & txtBonumber.Text & " saved successfully"
grdBackorders.Clear
cmbSuppliers.Text = ""
txtBonumber.Text = ""
txtAddress.Text = ""
grdBackorders.Rows = 1
frmBackorders.Tag = "Notsaveable"
cmdSupplier_Click


Case "Notsaveable"
Case ""
End Select

End Sub

Public Sub CreateNewOrder()
frmBackorders.Tag = "Notsaveable"


If cmbSuppliers.Text = "" Then
Load frmError
frmError.lblCap.Caption = "Please select a Supplier name first."
frmError.Show vbModal
Exit Sub
End If


Dim newline As Integer
ActiveReadServer "Select Max(Order_No) as Order_No from backorder_journal"
If rs.RecordCount > 0 Then
If IsNull(rs.Fields("Order_No")) Then
newline = 1
GoTo nulls
End If

newline = rs.Fields("Order_No") + 1
End If


nulls:

If newline < 1 Then newline = 1
rs.Close
txtBonumber.Text = newline
grdBackorders.Tag = "newbackorder"
grdBackorders.Clear
grdBackorders.Rows = 1
grdBackorders.Cols = 10
grdBackorders.TextMatrix(0, 0) = "User Number"
    grdBackorders.TextMatrix(0, 1) = "Location"
    grdBackorders.TextMatrix(0, 2) = "Supplier no"
    grdBackorders.TextMatrix(0, 3) = "Order no"
    grdBackorders.TextMatrix(0, 4) = "Product Code"
    grdBackorders.TextMatrix(0, 5) = "Description"
    grdBackorders.TextMatrix(0, 6) = "Quantity Ordered"
    grdBackorders.TextMatrix(0, 7) = "Price Ordered"
    grdBackorders.TextMatrix(0, 8) = "Order Date"
    grdBackorders.TextMatrix(0, 9) = "Expected Delivery_Date"
    grdBackorders.ColWidth(0) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(1) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(2) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(3) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(4) = grdBackorders.Width * 0.1
    grdBackorders.ColWidth(5) = grdBackorders.Width * 0.25
    grdBackorders.ColWidth(6) = grdBackorders.Width * 0.1
    grdBackorders.ColWidth(7) = grdBackorders.Width * 0.1
    grdBackorders.ColWidth(8) = grdBackorders.Width * 0.1
    grdBackorders.ColWidth(9) = grdBackorders.Width * 0.05
    grdBackorders.ColAlignment(0) = flexAlignLeftCenter
    grdBackorders.ColAlignment(1) = flexAlignLeftCenter
    grdBackorders.ColAlignment(2) = flexAlignLeftCenter
    grdBackorders.ColAlignment(3) = flexAlignLeftCenter
    grdBackorders.ColAlignment(4) = flexAlignRightCenter
    grdBackorders.ColAlignment(5) = flexAlignRightCenter
    grdBackorders.ColAlignment(6) = flexAlignRightCenter
    grdBackorders.ColAlignment(7) = flexAlignRightCenter
    grdBackorders.ColAlignment(8) = flexAlignRightCenter
    grdBackorders.ColAlignment(9) = flexAlignRightCenter

If grdBackorders.Tag = "newbackorder" Then
        
        
        grdBackorders.Rows = grdBackorders.Rows + 1
        i = grdBackorders.Rows - 1
        a = UserRecord.User_Number
        grdBackorders.TextMatrix(i, 0) = a
        grdBackorders.TextMatrix(i, 1) = "1"
        grdBackorders.TextMatrix(i, 2) = cmdSupplier.Tag
        grdBackorders.TextMatrix(i, 3) = txtBonumber.Text
        grdBackorders.TextMatrix(i, 8) = Date
        grdBackorders.Select i, 0
        frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
        End If

End Sub
Public Sub List_all_backorders(Supplier_number As String)
frmBackorders.Tag = "Notsaveable"
frmMain.stbBar.Panels(3) = "Records = 0"
grdBackorders.Clear
    grdBackorders.Rows = 1
    grdBackorders.Cols = 5
    grdBackorders.TextMatrix(0, 0) = "Backorder Number"
    grdBackorders.TextMatrix(0, 1) = "Supplier No"
    grdBackorders.TextMatrix(0, 2) = "Supplier Name"
    grdBackorders.TextMatrix(0, 3) = "Date"
    grdBackorders.TextMatrix(0, 4) = "Created by User"
    grdBackorders.ColWidth(0) = grdBackorders.Width * 0.2
    grdBackorders.ColWidth(1) = grdBackorders.Width * 0.2
    grdBackorders.ColWidth(2) = grdBackorders.Width * 0.2
    grdBackorders.ColWidth(3) = grdBackorders.Width * 0.2
    grdBackorders.ColWidth(4) = grdBackorders.Width * 0.2
    grdBackorders.ColAlignment(0) = flexAlignLeftCenter
    grdBackorders.ColAlignment(1) = flexAlignLeftCenter
    grdBackorders.ColAlignment(2) = flexAlignLeftCenter
    grdBackorders.ColAlignment(3) = flexAlignRightCenter
    grdBackorders.ColAlignment(4) = flexAlignRightCenter
ActiveReadServer " Select * from Backorder_view where supplier_no = " & "'" & _
Supplier_number & "'"
On Error GoTo trap
grdBackorders.Rows = rs.RecordCount + 1

    i = 0
    While Not rs.EOF
        i = i + 1
        grdBackorders.TextMatrix(i, 0) = rs.Fields("Order_No")
        grdBackorders.TextMatrix(i, 1) = rs.Fields("Supplier_No")
        grdBackorders.TextMatrix(i, 2) = rs.Fields("Supplier_Name")
        grdBackorders.TextMatrix(i, 3) = Format(rs.Fields("Date_time"), "DD/MM/YYYY")
        grdBackorders.TextMatrix(i, 4) = rs.Fields("First_Name")
        rs.MoveNext
    Wend
    frmMain.stbBar.Panels(3) = "Records = " & Val(rs.RecordCount)
    rs.Close
    If grdBackorders.Rows > 1 Then grdBackorders.Row = 1
trap:
End Sub
Public Sub Open_backordernumber(Backordernumber As String) ' by Backordernumber the actual order
frmMain.stbBar.Panels(3) = "Records = 0"
grdBackorders.Clear
grdBackorders.Rows = 1
grdBackorders.Cols = 10
grdBackorders.TextMatrix(0, 0) = "User Number"
    grdBackorders.TextMatrix(0, 1) = "Location"
    grdBackorders.TextMatrix(0, 2) = "Supplier no"
    grdBackorders.TextMatrix(0, 3) = "Order no"
    grdBackorders.TextMatrix(0, 4) = "Product Code"
    grdBackorders.TextMatrix(0, 5) = "Description"
    grdBackorders.TextMatrix(0, 6) = "Quantity Ordered"
    grdBackorders.TextMatrix(0, 7) = "Price Ordered"
    grdBackorders.TextMatrix(0, 8) = "Order Date"
    grdBackorders.TextMatrix(0, 9) = "Expected Delivery_Date"
    grdBackorders.ColWidth(0) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(1) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(2) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(3) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(4) = grdBackorders.Width * 0.1
    grdBackorders.ColWidth(5) = grdBackorders.Width * 0.2
    grdBackorders.ColWidth(6) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(7) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(8) = grdBackorders.Width * 0.05
    grdBackorders.ColWidth(9) = grdBackorders.Width * 0.05
    grdBackorders.ColAlignment(0) = flexAlignLeftCenter
    grdBackorders.ColAlignment(1) = flexAlignLeftCenter
    grdBackorders.ColAlignment(2) = flexAlignLeftCenter
    grdBackorders.ColAlignment(3) = flexAlignLeftCenter
    grdBackorders.ColAlignment(4) = flexAlignRightCenter
    grdBackorders.ColAlignment(5) = flexAlignRightCenter
    grdBackorders.ColAlignment(6) = flexAlignRightCenter
    grdBackorders.ColAlignment(7) = flexAlignRightCenter
    grdBackorders.ColAlignment(8) = flexAlignRightCenter
    grdBackorders.ColAlignment(9) = flexAlignRightCenter

ActiveReadServer " Select * from Backorder_journal where Order_no = " & "'" & _
Backordernumber & "'"
If Backordernumber = "Backorder Number" Then Exit Sub
grdBackorders.Rows = rs.RecordCount + 1

    i = 0
    While Not rs.EOF
        i = i + 1
        grdBackorders.TextMatrix(i, 0) = rs.Fields("User_No")
        grdBackorders.TextMatrix(i, 1) = rs.Fields("Location_No")
        grdBackorders.TextMatrix(i, 2) = rs.Fields("Supplier_No")
        grdBackorders.TextMatrix(i, 3) = rs.Fields("Order_no")
        grdBackorders.TextMatrix(i, 4) = rs.Fields("Product_Code")
        
        ActiveReadServer2 " Select * from Products where Product_Code = " & _
        "'" & rs.Fields("Product_Code") & "'"
        grdBackorders.TextMatrix(i, 5) = rs2.Fields("Description")
        rs2.Close
        grdBackorders.TextMatrix(i, 6) = rs.Fields("Qty_ordered")
        grdBackorders.TextMatrix(i, 7) = Format(rs.Fields("Price_ordered"), "0.00")
        grdBackorders.TextMatrix(i, 8) = Format(rs.Fields("Order_Date"), "DD/MM/YYYY")
        grdBackorders.TextMatrix(i, 9) = Format(rs.Fields("Delivery_Date"), "DD/MM/YYYY")
        rs.MoveNext
    Wend
    frmMain.stbBar.Panels(3) = "Records = " & Val(rs.RecordCount)
    rs.Close
    If grdBackorders.Rows > 1 Then grdBackorders.Row = 1
    
    
    ' enable  delete save - when edited (delete individual products or whole backorder)
        frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
    
    
    
    
    
    
    
    


End Sub
Public Sub Open_backorderdate()

End Sub

Private Sub ButtonEx1_Click()
Image6.Visible = True
mthViewStart.Visible = True
mthViewEnd.Visible = True
cmdOk.Visible = True
ButtonEx2.Visible = False
End Sub

Private Sub ButtonEx2_Click()
grdBackorders.Tag = "ByProduct"
cmbSuppliers.Text = ""
txtBonumber = ""
txtAddress = ""
Listall_products


End Sub

Private Sub cmbSuppliers_Change()

YYYY = frmBackorders.Tag
yy = grdBackorders.Tag
If grdBackorders.Tag = "Byorderno" And frmBackorders.Tag = "Notsaveable" Then
Exit Sub
End If
If cmbSuppliers.Text = "" Then Exit Sub


If cmbSuppliers.Text <> " " Then
Dim x, y, z
x = Mid(cmbSuppliers.Text, InStr(cmbSuppliers.Text, "("))
y = Len(Mid(cmbSuppliers.Text, InStr(cmbSuppliers.Text, "(")))
z = Left(cmbSuppliers.Text, Len(cmbSuppliers.Text) - y)
ActiveReadServer1 "Select * from Suppliers where Supplier_Name ='" & z & "'"
If rs1.EOF <> True Then
txtAddress = rs1.Fields("Address")
cmdSupplier.Tag = rs1.Fields("Supplier_No")
grdBackorders.Tag = "Byorderno"
End If
rs1.Close
End If
End Sub

Private Sub cmdOk_Click()
ButtonEx2.Visible = True
mthViewStart.Visible = False
mthViewEnd.Visible = False
Image6.Visible = False
cmdOk.Visible = False
lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & _
Format(mthViewEnd.Value, "DD MMM YYYY")

End Sub

Private Sub cmdSupplier_Click()
        frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
grdBackorders.Tag = "BySuppliers"
Listallsuppliers

'List_all_backorders (supplierdetail)
End Sub
Private Sub Listallsuppliers()
    frmMain.stbBar.Panels(3) = "Records = 0"
    grdBackorders.Cols = 5
    grdBackorders.TextMatrix(0, 0) = " Supplier Number"
    grdBackorders.TextMatrix(0, 1) = "Supplier Name"
    grdBackorders.TextMatrix(0, 2) = "Contact Person"
    grdBackorders.TextMatrix(0, 3) = "Tel.Number"
    grdBackorders.TextMatrix(0, 4) = "Balance "
    grdBackorders.ColWidth(0) = grdBackorders.Width * 0.2
    grdBackorders.ColWidth(1) = grdBackorders.Width * 0.3
    grdBackorders.ColWidth(2) = grdBackorders.Width * 0.22
    grdBackorders.ColWidth(3) = grdBackorders.Width * 0.15
    grdBackorders.ColWidth(4) = grdBackorders.Width * 0.13
    grdBackorders.ColAlignment(4) = flexAlignRightCenter
    grdBackorders.ColAlignment(0) = flexAlignLeftCenter
    grdBackorders.ColAlignment(1) = flexAlignLeftCenter
    grdBackorders.ColAlignment(2) = flexAlignLeftCenter
    grdBackorders.ColAlignment(3) = flexAlignLeftCenter
    grdBackorders.ColAlignment(4) = flexAlignRightCenter
    


'grdBackorders = 1
    ActiveReadServer "Select * from Suppliers order by Supplier_Name"
 
    grdBackorders.Rows = rs.RecordCount + 1
    i = 0
    While Not rs.EOF
        i = i + 1
        grdBackorders.TextMatrix(i, 0) = rs.Fields("Supplier_No")
        grdBackorders.TextMatrix(i, 1) = rs.Fields("Supplier_Name")
        grdBackorders.TextMatrix(i, 2) = rs.Fields("Contact_Person")
        grdBackorders.TextMatrix(i, 3) = rs.Fields("Business_Tel")
        grdBackorders.TextMatrix(i, 4) = Format(rs.Fields("Balance"), "0.00")
        rs.MoveNext
    Wend
    frmMain.stbBar.Panels(3) = "Records = " & Val(rs.RecordCount)
    rs.Close
    frmBackorders.Tag = "Byorderno"
    If grdBackorders.Rows > 1 Then grdBackorders.Row = 1
End Sub



Private Sub Form_Initialize()
frmMain.cmdMenu(4).Tag = "Backorders"
frmBackorders.WindowState = 2
frmMain.stbBar.Panels(3) = "Records = 0"
frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmBackorders.Tag = "Notsaveable"
ActiveReadServer " Select * from Suppliers"
While Not rs.EOF
cmbSuppliers.AddItem (rs.Fields("Supplier_Name")) & " (" & rs.Fields("Supplier_No") & ")"
rs.MoveNext
Wend
rs.Close

mthViewStart.Value = Date
mthViewEnd.Value = Date
lblDate.Caption = Format(mthViewStart.Value, "DD MMM YYYY") & " to " & _
Format(mthViewEnd.Value, "DD MMM YYYY")

End Sub

Private Sub Form_Load()
frmBackorders.Left = 0
frmBackorders.top = 0

End Sub



Private Sub grdBackorders_DblClick()
    
    
    Select Case grdBackorders.Tag
    
    Case "BySuppliers"
    ' Select by Supplier
    
    frmBackorders.Tag = "Notsaveable"
    If grdBackorders.TextMatrix(0, 0) = " Supplier Number" Then
    frmBackorders.Tag = "Notsaveable"
    cmbSuppliers.Text = grdBackorders.TextMatrix(grdBackorders.Row, 1) & _
    " (" & grdBackorders.TextMatrix(grdBackorders.Row, 0) & ")"
    Comboproducts.Text = " "
    
    ActiveReadServer1 "Select * from Suppliers where Supplier_No='" & _
    (grdBackorders.TextMatrix(grdBackorders.Row, 0)) & "'"
    
    txtAddress = rs1.Fields("Address")
    cmdSupplier.Tag = rs1.Fields("Supplier_No")
    List_all_backorders (grdBackorders.TextMatrix(grdBackorders.Row, 0))
    ButtonEx1.Visible = True
    grdBackorders.Tag = "Byorderno"
    End If
    If grdBackorders.Rows > 1 Then grdBackorders.Row = 1
    
 
    Case "Byorderno" ' Select by backorder no
        frmBackorders.Tag = "Notsaveable"
        If grdBackorders.TextMatrix(0, 0) = "Backorder Number" Then
        ButtonEx1.Visible = False
        txtBonumber = (grdBackorders.TextMatrix(grdBackorders.Row, 0))
        frmBackorders.Tag = "Suppliers"
        Open_backordernumber (grdBackorders.TextMatrix(grdBackorders.Row, 0))
        If grdBackorders.Rows > 1 Then grdBackorders.Row = 1
        End If
      


    Case "ByProduct" ' Select by Product
    frmBackorders.Tag = "Notsaveable"
    If grdBackorders.TextMatrix(0, 0) = " Product Number" Then
    'Select By product no
    Comboproducts.Text = (grdBackorders.TextMatrix(grdBackorders.Row, 1))
    cmbSuppliers.Text = " "
    txtAddress = " "
    frmBackorders.Tag = "Products"
    Select_by_Product (grdBackorders.TextMatrix(grdBackorders.Row, 0))
    ButtonEx1.Visible = True
    If grdBackorders.Rows > 1 Then grdBackorders.Row = 1
    End If
    
    
End Select


End Sub

Private Sub Select_by_Product(Productno_code As String)
frmMain.stbBar.Panels(3) = "Records = 0"
grdBackorders.Clear
grdBackorders.Rows = 1
grdBackorders.Cols = 6

frmBackorders.Tag = "Notsaveable"
    grdBackorders.TextMatrix(0, 0) = "Backorder Number"
    grdBackorders.TextMatrix(0, 1) = "Supplier No"
    grdBackorders.TextMatrix(0, 2) = "Order Date"
    grdBackorders.TextMatrix(0, 3) = "Expected Date"
    grdBackorders.TextMatrix(0, 4) = "Total Value"
    grdBackorders.TextMatrix(0, 5) = "Created by User"
    grdBackorders.ColWidth(0) = grdBackorders.Width * 0.166
    grdBackorders.ColWidth(1) = grdBackorders.Width * 0.166
    grdBackorders.ColWidth(2) = grdBackorders.Width * 0.166
    grdBackorders.ColWidth(3) = grdBackorders.Width * 0.166
    grdBackorders.ColWidth(4) = grdBackorders.Width * 0.166
    grdBackorders.ColWidth(5) = grdBackorders.Width * 0.166
    grdBackorders.ColAlignment(0) = flexAlignLeftCenter
    grdBackorders.ColAlignment(1) = flexAlignLeftCenter
    grdBackorders.ColAlignment(2) = flexAlignLeftCenter
    grdBackorders.ColAlignment(3) = flexAlignLeftCenter
    grdBackorders.ColAlignment(4) = flexAlignRightCenter
    grdBackorders.ColAlignment(5) = flexAlignLeftCenter
ActiveReadServer " Select * from Backorder_journal where Product_Code = " & _
"'" & Productno_code & "'"
On Error GoTo trap
grdBackorders.Rows = rs.RecordCount + 1

    i = 0
    While Not rs.EOF
        i = i + 1
        grdBackorders.TextMatrix(i, 0) = rs.Fields("Order_No")
        grdBackorders.TextMatrix(i, 1) = rs.Fields("Supplier_No")
        grdBackorders.TextMatrix(i, 2) = rs.Fields("Order_Date")
        grdBackorders.TextMatrix(i, 3) = rs.Fields("Delivery_Date")
        grdBackorders.TextMatrix(i, 4) = Format(rs.Fields("Line_Total"), "0.00")
        
        ActiveReadServer2 " Select * from Backorder_journal where Order_No = " & _
        "'" & rs.Fields("Order_No") & "'"
        grdBackorders.TextMatrix(i, 5) = rs2.Fields("User_No")
        rs2.Close
        rs.MoveNext
    Wend
    frmMain.stbBar.Panels(3) = "Records = " & Val(rs.RecordCount)
    rs.Close
    If grdBackorders.Rows > 1 Then grdBackorders.Row = 1
    
        ActiveReadServer " Select * from Suppliers where  Supplier_No = '" & Trim(grdBackorders.TextMatrix(1, 1)) & "'"
        cmbSuppliers.Text = rs.Fields("Supplier_No") & "-" & rs.Fields("Supplier_Name")
        txtAddress.Text = rs.Fields("Address")
frmBackorders.Tag = "BySuppliers"
grdBackorders.Tag = "Byorderno"

trap:

End Sub

Private Sub Listall_products()
    frmMain.stbBar.Panels(3) = "Records = 0"
    frmBackorders.Tag = "Notsaveable"
    grdBackorders.Cols = 5
    grdBackorders.TextMatrix(0, 0) = " Product Number"
    grdBackorders.TextMatrix(0, 1) = "Product Description"
    grdBackorders.TextMatrix(0, 2) = "Average Cost"
    grdBackorders.TextMatrix(0, 3) = "Landed Cost"
    grdBackorders.TextMatrix(0, 4) = "Selling Price"
    grdBackorders.ColWidth(0) = grdBackorders.Width * 0.2
    grdBackorders.ColWidth(1) = grdBackorders.Width * 0.3
    grdBackorders.ColWidth(2) = grdBackorders.Width * 0.22
    grdBackorders.ColWidth(3) = grdBackorders.Width * 0.15
    grdBackorders.ColWidth(4) = grdBackorders.Width * 0.13
    grdBackorders.ColAlignment(4) = flexAlignRightCenter
    grdBackorders.ColAlignment(0) = flexAlignLeftCenter
    grdBackorders.ColAlignment(1) = flexAlignLeftCenter
    grdBackorders.ColAlignment(2) = flexAlignLeftCenter
    grdBackorders.ColAlignment(3) = flexAlignLeftCenter
    grdBackorders.ColAlignment(4) = flexAlignRightCenter
    
    ActiveReadServer "Select * from Products where Stock_Item = '1' order by Description "
    grdBackorders.Rows = rs.RecordCount + 1
    i = 0
    While Not rs.EOF
        i = i + 1
        grdBackorders.TextMatrix(i, 0) = rs.Fields("Product_Code")
        grdBackorders.TextMatrix(i, 1) = rs.Fields("Description")
        grdBackorders.TextMatrix(i, 2) = rs.Fields("Ave_Cost")
        grdBackorders.TextMatrix(i, 3) = rs.Fields("Landed_Cost")
        grdBackorders.TextMatrix(i, 4) = Format(rs.Fields("Selling_Price"), "0.00")
        rs.MoveNext
    Wend
    frmMain.stbBar.Panels(3) = "Records = " & Val(rs.RecordCount)
    
    rs.Close
    If grdBackorders.Rows > 1 Then grdBackorders.Row = 1
    
    
End Sub

Private Sub grdBackorders_KeyDown(KeyCode As Integer, Shift As Integer)

Dim i As Integer
Dim a
Dim Stringtotallength As Integer, Codelength As Integer, Descriptionlength As Integer
                        Dim Descriptinstr As String
Select Case KeyCode
Case 13 ' Enter button
        If grdBackorders.Rows = 1 Then Exit Sub
        '
        'if new order has been selected
        If grdBackorders.Col <> 5 Then
        If grdBackorders.Tag = "newbackorder" Then
        
        
        grdBackorders.Rows = grdBackorders.Rows + 1
        i = grdBackorders.Rows - 1
        a = UserRecord.User_Number
        grdBackorders.TextMatrix(i, 0) = a
        grdBackorders.TextMatrix(i, 1) = "1"
        grdBackorders.TextMatrix(i, 2) = cmdSupplier.Tag
        grdBackorders.TextMatrix(i, 3) = txtBonumber.Text
        grdBackorders.TextMatrix(i, 8) = Date
        grdBackorders.Select i, 0
        frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmBackorders.Tag = "Notsaveable"
        End If
        End If
        
        
        If grdBackorders.Col = 5 And grdBackorders.Rows > 1 Then
        frmSearch.Tag = "Back_Order"
        frmSearch.Show vbModal
        
                    Select Case frmSearch.Tag
                    Case ""
                    Case Else
                        DoEvents
                        
                        Descriptinstr = ""
                        grdBackorders.TextMatrix(grdBackorders.Row, 4) = Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1))
                        Stringtotallength = Len(frmSearch.Tag)
                        Codelength = Len(Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1)))
                        Codelength = Codelength + 3
                        Descriptionlength = Stringtotallength - (Codelength)
                        Descriptinstr = Left(frmSearch.Tag, Descriptionlength)
                        Descriptinstr = Trim(Descriptinstr)
                        grdBackorders.TextMatrix(grdBackorders.Row, 5) = Descriptinstr
                        grdBackorders.TextMatrix(grdBackorders.Row, 6) = "1"
                        ActiveReadServer " Select Landed_Cost from Products as Landed_Cost where Product_Code = '" & grdBackorders.TextMatrix(grdBackorders.Row, 4) & "'"
                        If rs.RecordCount > 0 Then
                        grdBackorders.TextMatrix(grdBackorders.Row, 7) = rs.Fields("Landed_cost")
                        End If
                        grdBackorders.TextMatrix(grdBackorders.Row, 9) = Date
                        frmBackorders.Tag = "Saveable"
                        End Select
            End If
        
        
        


Case 46 ' Delete button

 If grdBackorders.Rows = 1 Then frmBackorders.Tag = "NotSaveable": Exit Sub
 If grdBackorders.Rows > 1 Then
 grdBackorders.RemoveItem (grdBackorders.Row)
 End If
            'Calculate_Totals
            If grdBackorders.Rows = 1 Then
                frmMain.Toolbar1.Buttons(2).Enabled = False
                frmMain.Toolbar1.Buttons(3).Enabled = True
                frmMain.Toolbar1.Buttons(4).Enabled = False
            End If




Case 40 ' Down Arrow


        If grdBackorders.Tag = "newbackorder" Then
        
        If grdBackorders.Row = grdBackorders.Rows - 1 Then
        grdBackorders.Rows = grdBackorders.Rows + 1
        i = grdBackorders.Rows - 1
        a = UserRecord.User_Number
        grdBackorders.TextMatrix(i, 0) = a

        grdBackorders.TextMatrix(i, 1) = "1"
        grdBackorders.TextMatrix(i, 2) = cmdSupplier.Tag
        grdBackorders.TextMatrix(i, 3) = txtBonumber.Text
        grdBackorders.TextMatrix(i, 8) = Date
        grdBackorders.Select i, 0
        frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
        frmMain.Toolbar1.Buttons(2).Caption = "New"
        frmMain.Toolbar1.Buttons(2).Enabled = True
        frmMain.Toolbar1.Buttons(3).Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = True
        grdBackorders.Row = grdBackorders.Rows - 1
        grdBackorders.Col = grdBackorders.Cols - 5
        frmBackorders.Tag = "Notsaveable"
        End If
        End If
        
End Select

End Sub

Private Sub grdBackorders_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    
    Select Case Col

        Case 0
            Cancel = True
        Case 1
        Cancel = True
        Case 2
             Cancel = True
        Case 3
             Cancel = True
        Case 4, 5
Cancel = True

        Case 6
        Case 7
        Case 8
            Cancel = True
        Case 9
        End Select

End Sub
Private Sub test()
'Dim locname As String, length, strings As String, locno As String
            'strings = grdBackorders.TextMatrix(i, 1)
            'length = Val(Len(grdBackorders.TextMatrix(i, 1)))
            'First = Val(LTrim(InStr(grdBackorders.TextMatrix(i, 1), "-") + 2))
            'Seconds = Val(length) - (Val(First)) + 1
                'If InStr(strings, "-") > 0 Then
                'locname = Mid(strings, First, Seconds)
               'End If
               
               
               
               
               
               'With grdOrder
'    Select Case KeyCode
'        Case 13 'Enter
'            If grdOrder.Col = 1 Then
'                Load frmSearch
'                frmSearch.Tag = "Back_Order"
'                frmSearch.Show vbModal
'                Select Case frmSearch.Tag
'                    Case ""
'                    Case Else
'                        grdOrder.TextMatrix(grdOrder.Row, 2) = frmSearch.Tag
'                        grdOrder.TextMatrix(grdOrder.Row, 3) = "1"
'                        grdOrder.TextMatrix(grdOrder.Row, 4) = "0"
'                        grdOrder.TextMatrix(grdOrder.Row, 5) = "0"
'                        grdOrder.TextMatrix(grdOrder.Row, 8) = "0.00"
'                        grdOrder.TextMatrix(grdOrder.Row, 0) = Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1))
'                        ActiveReadServer "Select Pack_Size,Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
'                        If rs.RecordCount > 0 Then
'                            grdOrder.TextMatrix(grdOrder.Row, 6) = rs.Fields("Landed_Cost")
'                            grdOrder.TextMatrix(grdOrder.Row, 7) = rs.Fields("Sales_Tax") & "%"
'                            grdOrder.TextMatrix(grdOrder.Row, 9) = rs.Fields("Department_No")
'                            grdOrder.TextMatrix(grdOrder.Row, 3) = rs.Fields("Pack_Size")
'                        End If
'                        rs.Close
'                        grdOrder.Col = 5
'                        ActiveReadServer2 "Select * from Suppliers_Links_View where Supplier_No = '" & lblGRV.Caption & "' and Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
'                        If rs2.RecordCount > 0 Then
'                            grdOrder.TextMatrix(grdOrder.Row, 1) = rs2.Fields("Supplier_Code")
'                        End If
'                        rs2.Close
'                        frmMain.Toolbar1.Buttons(2).Enabled = True
'                        frmMain.Toolbar1.Buttons(3).Enabled = False
'                        frmMain.Toolbar1.Buttons(4).Enabled = True
'                        frmSearch.Tag = ""
'                End Select
'            End If
'        Case 106
'            frmProdSearch.Show vbModal
'            If frmOrder.Tag <> "" Then
'                grdOrder.TextMatrix(grdOrder.Row, 1) = frmOrder.Tag
'                frmOrder.Tag = ""
'                grdOrder.TextMatrix(grdOrder.Row, 3) = "1"
'                grdOrder.TextMatrix(grdOrder.Row, 4) = "0"
'                grdOrder.TextMatrix(grdOrder.Row, 5) = "0"
'                grdOrder.TextMatrix(grdOrder.Row, 0) = Trim(Mid(grdOrder.Text, InStrRev(grdOrder.Text, "-") + 1))
'                ActiveReadServer "Select Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
'                If rs.RecordCount > 0 Then
'                    grdOrder.TextMatrix(grdOrder.Row, 6) = rs.Fields("Landed_Cost")
'                    grdOrder.TextMatrix(grdOrder.Row, 7) = rs.Fields("Sales_Tax") & "%"
'                    grdOrder.TextMatrix(grdOrder.Row, 9) = rs.Fields("Department_No")
'                End If
'                rs.Close
'                grdOrder.Col = 5
'                frmMain.Toolbar1.Buttons(2).Enabled = True
'                frmMain.Toolbar1.Buttons(3).Enabled = False
'                frmMain.Toolbar1.Buttons(4).Enabled = True
'            End If
'        Case 46
'            grdOrder.RemoveItem (grdOrder.Row)
'            Calculate_Totals
'            If .Rows = 1 Then
'                frmMain.Toolbar1.Buttons(2).Enabled = False
'                frmMain.Toolbar1.Buttons(3).Enabled = True
'                frmMain.Toolbar1.Buttons(4).Enabled = False
'            End If
'        Case 45, 48 To 57, 96 To 105, 109, 110, 189
'            Select Case grdOrder.Col
'                Case 2, 3, 4, 7
'                Case Else
'                    .EditCell
'            End Select
'        Case 37 'left
'            If .Col = 1 Then
'                KeyCode = 0
'                .Col = 9
'            End If
'        Case 38 'up
'            If grdOrder.Row = 1 Then
'                cmbSuppliers.SetFocus
'            End If
'            If Trim(grdOrder.TextMatrix(grdOrder.Row, 2)) = "" Then
'                grdOrder.RemoveItem (grdOrder.Row)
'            End If
'        Case 39 'Right
'            If .Col = 9 Then
'                KeyCode = 0
'                .Col = 1
'            End If
'        Case 40 'down
'            If Trim(grdOrder.TextMatrix(grdOrder.Row, 2)) = "" Then
'                KeyCode = 0
'                cmbSuppliers.SetFocus
'                If grdOrder.Rows > 2 Then grdOrder.RemoveItem (grdOrder.Row)
'            Else
'                If .Row = .Rows - 1 Then
'                    .Rows = .Rows + 1
'                    .Row = .Rows - 1
'                    grdOrder.ShowCell .Row, 0
'                    .Col = 1
'                End If
'            End If
'    End Select
'End With
               
'        If grdBackorders.ColComboList(1) = "" Then
'        ActiveReadServer "Select * from Locations"
'        While Not rs.EOF
'        grdBackorders.ColComboList(1) = grdBackorders.ColComboList(1) & _
'        rs.Fields("Location_No") & "- " & rs.Fields("Loc_Name") & "|"
'        rs.MoveNext
'        Wend
'        rs.Close
'        End If
        
'        ActiveReadServer "Select * from Products"
'        While Not rs.EOF
'        grdBackorders.ColComboList(5) = grdBackorders.ColComboList(5) & _
'        rs.Fields("Product_Code") & "- " & rs.Fields("Description") & "|"
'        rs.MoveNext
'        Wend
'        rs.Close


 '???????????????????????????????????????????????????????????
                        'grdBackorders.TextMatrix(grdBackorders.Row, 5) = Left(frmSearch.Tag, Descriptionlength)
                        
                        
'                        grdOrder.TextMatrix(grdOrder.Row, 8) = "0.00"
'                        grdOrder.TextMatrix(grdOrder.Row, 0) = Trim(Mid(frmSearch.Tag, InStrRev(frmSearch.Tag, "-") + 1))
'                        ActiveReadServer "Select Pack_Size,Ave_Cost, Landed_Cost,Sales_Tax,Department_No from Products where Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
'                        If rs.RecordCount > 0 Then
'                            grdOrder.TextMatrix(grdOrder.Row, 6) = rs.Fields("Landed_Cost")
'                            grdOrder.TextMatrix(grdOrder.Row, 7) = rs.Fields("Sales_Tax") & "%"
'                            grdOrder.TextMatrix(grdOrder.Row, 9) = rs.Fields("Department_No")
'                            grdOrder.TextMatrix(grdOrder.Row, 3) = rs.Fields("Pack_Size")
'                        End If
'                        rs.Close
'                        grdOrder.Col = 5
'                        ActiveReadServer2 "Select * from Suppliers_Links_View where Supplier_No = '" & lblGRV.Caption & "' and Product_Code = '" & grdOrder.TextMatrix(grdOrder.Row, 0) & "'"
'                        If rs2.RecordCount > 0 Then
'                            grdOrder.TextMatrix(grdOrder.Row, 1) = rs2.Fields("Supplier_Code")
'                        End If
'                        rs2.Close
'                        frmMain.Toolbar1.Buttons(2).Enabled = True
'                        frmMain.Toolbar1.Buttons(3).Enabled = False
'                        frmMain.Toolbar1.Buttons(4).Enabled = True
'                        frmSearch.Tag = ""



'grdBackorders.TextMatrix(0, 1) = "Location"
'    grdBackorders.TextMatrix(0, 2) = "Supplier no"
'    grdBackorders.TextMatrix(0, 3) = "Grv no"
'    grdBackorders.TextMatrix(0, 4) = "Product Code"
'    grdBackorders.TextMatrix(0, 5) = "Description"
'    grdBackorders.TextMatrix(0, 6) = "Quantity Ordered"
'    grdBackorders.TextMatrix(0, 7) = "Price Ordered"
'    grdBackorders.TextMatrix(0, 8) = "Order Date"
'    grdBackorders.TextMatrix(0, 9) = "Delivery_Date"


End Sub
