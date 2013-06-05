VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmProd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Products..."
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   Icon            =   "frmProd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProducts 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      HasDC           =   0   'False
      Height          =   9465
      Index           =   0
      Left            =   0
      ScaleHeight     =   9465
      ScaleWidth      =   14715
      TabIndex        =   14
      Top             =   0
      Width           =   14715
      Begin VB.PictureBox picTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3735
         Index           =   0
         Left            =   5160
         ScaleHeight     =   3705
         ScaleWidth      =   3315
         TabIndex        =   18
         Top             =   1680
         Width           =   3345
         Begin VB.PictureBox picBox 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3510
            Index           =   0
            Left            =   120
            ScaleHeight     =   3510
            ScaleWidth      =   4185
            TabIndex        =   32
            Top             =   120
            Width           =   4185
            Begin VB.TextBox txtSupp 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   1260
               TabIndex        =   12
               Top             =   2895
               Width           =   1785
            End
            Begin VB.TextBox txtPackSize 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   1260
               TabIndex        =   13
               Text            =   "1"
               Top             =   3255
               Width           =   1155
            End
            Begin VB.TextBox txtEmpty 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   4770
               TabIndex        =   41
               Top             =   2085
               Width           =   1155
            End
            Begin VB.CheckBox chkWhole 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Manage Stock as Whole Units Only"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   40
               Top             =   2385
               Width           =   3135
            End
            Begin VB.CheckBox chkScale 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Scale Item"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   39
               Top             =   1680
               Width           =   2385
            End
            Begin VB.CheckBox chkDelete 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Delete when Stock runs out"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   38
               Top             =   2040
               Width           =   2385
            End
            Begin VB.CheckBox chkTouch 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Display on Touch"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   37
               Top             =   990
               Width           =   1875
            End
            Begin VB.CheckBox chkRecipe 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Recipe Item"
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   36
               Top             =   645
               Width           =   1875
            End
            Begin VB.CheckBox chkSales 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Sales Item"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   35
               Top             =   285
               Value           =   1  'Checked
               Width           =   1875
            End
            Begin VB.CheckBox chkStock 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Stock Keeping Item"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   34
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1875
            End
            Begin VB.CheckBox chkDeposit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Deposit Bearing Item"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   33
               Top             =   1335
               Width           =   2385
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   3180
               Y1              =   2730
               Y2              =   2730
            End
            Begin MSForms.Image Image2 
               Height          =   285
               Left            =   1110
               Top             =   2850
               Width           =   1995
               BackColor       =   16777215
               Size            =   "3519;503"
            End
            Begin MSForms.Label Label2 
               Height          =   225
               Index           =   1
               Left            =   -510
               TabIndex        =   43
               Top             =   2895
               Width           =   1545
               BackColor       =   -2147483643
               Caption         =   "Supplier Code:"
               Size            =   "2725;397"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.Label Label2 
               Height          =   225
               Index           =   17
               Left            =   -510
               TabIndex        =   42
               Top             =   3255
               Width           =   1545
               BackColor       =   -2147483643
               Caption         =   "Pack Size:"
               Size            =   "2725;397"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.Image Image12 
               Height          =   285
               Left            =   1110
               Top             =   3210
               Width           =   1365
               BackColor       =   16777215
               Size            =   "2408;503"
            End
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   3330
            Y1              =   2850
            Y2              =   2850
         End
      End
      Begin VB.TextBox txtLandCost 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2860
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   3390
         Width           =   1725
      End
      Begin VB.TextBox txtMarkup 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2860
         TabIndex        =   6
         Text            =   "0"
         Top             =   3745
         Width           =   1725
      End
      Begin VB.TextBox txtGross 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   2860
         TabIndex        =   7
         Text            =   "0"
         Top             =   4130
         Width           =   1725
      End
      Begin VB.TextBox txtSellExcl 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2860
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   4490
         Width           =   1725
      End
      Begin VB.TextBox txtSellIncl 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2860
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   5170
         Width           =   1725
      End
      Begin VB.TextBox txtProductCode 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2905
         TabIndex        =   10
         Top             =   1030
         Width           =   1785
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2905
         TabIndex        =   0
         Top             =   1375
         Width           =   5535
      End
      Begin VB.TextBox txtShort 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2905
         TabIndex        =   1
         Top             =   1725
         Width           =   2115
      End
      Begin VB.TextBox txtUnitSize 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2905
         TabIndex        =   3
         Top             =   2425
         Width           =   2115
      End
      Begin VB.PictureBox PicBC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4740
         ScaleHeight     =   255
         ScaleWidth      =   315
         TabIndex        =   15
         Top             =   990
         Visible         =   0   'False
         Width           =   345
         Begin VB.Label lblCode 
            BackColor       =   &H00C0FFC0&
            Caption         =   "BC"
            Height          =   165
            Left            =   60
            TabIndex        =   16
            Top             =   30
            Width           =   285
         End
      End
      Begin btButtonEx.ButtonEx cmdBarcode 
         Height          =   285
         Left            =   5190
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   990
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "Change Barcode..."
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignHorz=   0
      End
      Begin btButtonEx.ButtonEx cmdCancel 
         Height          =   345
         Left            =   6990
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
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
      Begin btButtonEx.ButtonEx cmdSave 
         Height          =   345
         Left            =   5400
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "Save"
         Enabled         =   0   'False
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
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   14
         Left            =   1260
         TabIndex        =   45
         Top             =   2760
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Department:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   8
         Left            =   1260
         TabIndex        =   44
         Top             =   5160
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Selling Price (incl):"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   0
         Left            =   2760
         Top             =   3690
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   12
         Left            =   1260
         TabIndex        =   31
         Top             =   1395
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
         Left            =   1260
         TabIndex        =   30
         Top             =   1740
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
         Left            =   1260
         TabIndex        =   29
         Top             =   2085
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Unit of Measure:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1320
         X2              =   8490
         Y1              =   750
         Y2              =   750
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   90
         TabIndex        =   28
         Top             =   690
         Width           =   1185
         ForeColor       =   12582912
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "General"
         Size            =   "2090;450"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1320
         X2              =   5100
         Y1              =   3210
         Y2              =   3210
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   27
         Top             =   3180
         Width           =   1185
         ForeColor       =   12582912
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Pricing"
         Size            =   "2090;450"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbUnit 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Tag             =   "Up"
         Top             =   2040
         Width           =   2325
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4101;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbDepartments 
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Tag             =   "Up"
         Top             =   2730
         Width           =   2325
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4101;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbTax 
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Tag             =   "Up"
         Top             =   4770
         Width           =   2325
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4101;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   6
         Left            =   1260
         TabIndex        =   26
         Top             =   1050
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
         Left            =   1260
         TabIndex        =   25
         Top             =   4800
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Tax:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   4
         Left            =   1260
         TabIndex        =   24
         Top             =   4455
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Selling Price (excl):"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   3
         Left            =   1260
         TabIndex        =   23
         Top             =   4110
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Gross Profit%:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   1260
         TabIndex        =   22
         Top             =   3765
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Markup%"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblCost 
         Height          =   225
         Left            =   1260
         TabIndex        =   21
         Top             =   3420
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
         Left            =   1260
         TabIndex        =   20
         Top             =   2400
         Width           =   1395
         BackColor       =   -2147483643
         Caption         =   "Unit Size:"
         Size            =   "2461;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   900
         TabIndex        =   19
         Top             =   240
         Width           =   3105
         ForeColor       =   0
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Product Details"
         Size            =   "5477;450"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Image Image1 
         Height          =   345
         Left            =   840
         Top             =   210
         Width           =   3195
         BackColor       =   16777215
         Size            =   "5636;609"
         VariousPropertyBits=   19
      End
      Begin MSForms.Image Image6 
         Height          =   300
         Left            =   2760
         Top             =   3330
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   1
         Left            =   2760
         Top             =   4050
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   2
         Left            =   2760
         Top             =   4410
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   3
         Left            =   2760
         Top             =   3330
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image5 
         Height          =   300
         Index           =   4
         Left            =   2760
         Top             =   5110
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;529"
      End
      Begin MSForms.Image Image7 
         Height          =   285
         Index           =   0
         Left            =   2760
         Top             =   990
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;503"
      End
      Begin MSForms.Image Image8 
         Height          =   285
         Left            =   2760
         Top             =   1330
         Width           =   5745
         BackColor       =   16777215
         Size            =   "10134;503"
      End
      Begin MSForms.Image Image9 
         Height          =   285
         Left            =   2760
         Top             =   1680
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;503"
      End
      Begin MSForms.Image Image7 
         Height          =   285
         Index           =   1
         Left            =   2760
         Top             =   2380
         Width           =   2325
         BackColor       =   16777215
         Size            =   "4101;503"
      End
   End
End
Attribute VB_Name = "frmProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
    With frmProd
        If .cmbDepartments.Text = "<Unbound>" Then
            Depart = "0"
        Else
            Depart = Trim(Mid(.cmbDepartments.Text, 1, InStr(.cmbDepartments.Text, " -")))
        End If
        TaxRate = Mid(.cmbTax.Text, InStr(.cmbTax.Text, "-") + 2, InStr(.cmbTax.Text, "%") - InStr(.cmbTax.Text, "-") - 2)
        TaxType = Mid(.cmbTax.Text, 1, InStr(.cmbTax.Text, "-") - 2)
        ActiveReadServer "Select product_code from Products where Product_code='" & .txtProductCode.Text & "'"
        Select Case rs.RecordCount
            Case 0
                ActiveUpdateServer "INSERT INTO [Products]([Product_Code], [Description], [Short_Description]," & _
                "[Department_No], [Pack_Size], [Unit_Size], [Unit_of_Measure], [Maximum_Discount], [Sales_Item]," & _
                "[Stock_Item], [Returnable_Item], [Recipe_Item], [Touch_Item], [Scale_Item], [Landed_Cost]," & _
                "[Ave_Cost], [Selling_Price], [Sales_Tax], [Tax_Type], [Once_off], [Date_Created], [Date_Updated])" & _
                "VALUES('" & .txtProductCode.Text & "','" & .txtDescription.Text & "','" & .txtShort.Text & "','" & Depart & "'" & _
                ",'" & .txtPackSize.Text & "','" & .txtUnitSize.Text & "','" & .cmbUnit.Text & "',0," & .chkSales.Value & _
                "," & .chkStock.Value & "," & .chkDeposit.Value & "," & .chkRecipe.Value & "," & .chkTouch.Value & "," & .chkScale.Value & "," & Val(.txtLandCost.Text) & _
                "," & Val(.txtLandCost.Text) & "," & Val(.txtSellIncl.Text) & "," & TaxRate & "," & TaxType & "," & .chkDelete.Value & ",Getdate(),Getdate())"
            Case 1
                If Val(Trim(Mid(.txtLandCost.ToolTipText, InStr(.txtLandCost.ToolTipText, ":") + 1))) = 0 Then
                    AveCost = Val(.txtLandCost.Text)
                    .txtLandCost.ToolTipText = " Average Cost: " & Format(Val(.txtLandCost.Text), "0.00") & " "
                Else
                    ActiveReadServer1 "Select Sum(Stock_on_Hand) as Stock_on_Hand from Quantities where Product_Code = '" & .txtProductCode.Text & "'"
                    If rs1.RecordCount > 0 Then
                        If Val(rs1.Fields("Stock_on_Hand") & "") = 0 Then
                            AveCost = .txtLandCost.Text
                        Else
                             AveCost = Val(Trim(Mid(.txtLandCost.ToolTipText, InStr(.txtLandCost.ToolTipText, ":") + 1)))
                        End If
                    End If
                    rs1.Close
                End If
                  ActiveUpdateServer "UPDATE [Products] SET " & _
                "[Description]='" & .txtDescription.Text & "'," & _
                "[Short_Description]='" & .txtShort.Text & "'," & _
                "[Department_No]='" & Depart & "'," & _
                "[Pack_Size]='" & .txtPackSize.Text & "'," & _
                "[Unit_Size]='" & .txtUnitSize.Text & "'," & _
                "[Unit_of_Measure]='" & .cmbUnit.Text & "'," & _
                "[Maximum_Discount]=0," & _
                "[Sales_Item]=" & .chkSales.Value & "," & _
                "[Stock_Item]=" & .chkStock.Value & "," & _
                "[Returnable_Item]=" & .chkDeposit.Value & "," & _
                "[Recipe_Item]=" & .chkRecipe.Value & "," & _
                "[Touch_Item]=" & .chkTouch.Value & "," & _
                "[Scale_Item]=" & .chkScale.Value & "," & _
                "[Whole_Unit]=" & .chkWhole.Value & "," & _
                "[Landed_Cost]=" & Val(.txtLandCost.Text) & "," & _
                "[Ave_Cost]=" & AveCost & "," & _
                "[Selling_Price]=" & .txtSellIncl.Text & "," & _
                "[Sales_Tax]=" & TaxRate & "," & _
                "[Tax_Type]=" & TaxType & "," & _
                "[Once_off]=" & .chkDelete.Value & ",[Date_Updated]=Getdate(),[Weight_Empty]='" & Val(.txtEmpty.Text) & "'" & _
                " WHERE Product_Code='" & .txtProductCode.Text & "'"
        End Select
        rs.Close
        Select Case frmProd.Caption
            Case "Create a Product..."
                frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0) = txtProductCode.Text
            Case "Edit a Product..."
                frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0) = txtProductCode.Text
        End Select
        ActiveReadServer2 "Select * from Supplier_Links where Product_Code = '" & txtProductCode.Text & "' and Supplier_No = '" & Replace(Mid(frmGRV.cmbSuppliers, InStrRev(frmGRV.cmbSuppliers, "(") + 1), ")", "") & "'"
        If rs2.RecordCount = 0 Then
            ActiveUpdateServer "INSERT INTO [Supplier_Links]([Supplier_No], [Product_Code], [Supplier_Code], [List_Price], [Date_Time])" & _
            " VALUES('" & Replace(Mid(frmGRV.cmbSuppliers, InStrRev(frmGRV.cmbSuppliers, "(") + 1), ")", "") & "', '" & txtProductCode.Text & "', '" & txtSupp.Text & "'," & frmGRV.grdGRV.ValueMatrix(i, 7) & ",getdate())"
        Else
            ActiveUpdateServer "Update [Supplier_Links]" & _
            " SET [Supplier_No]='" & Replace(Mid(frmGRV.cmbSuppliers, InStrRev(frmGRV.cmbSuppliers, "(") + 1), ")", "") & "', [Product_Code]='" & txtProductCode.Text & "'" & _
            ", [List_Price]= " & frmGRV.grdGRV.ValueMatrix(i, 7) & ", [Date_Time]= getdate(),Supplier_Code = '" & txtSupp.Text & "'" & _
            " WHERE Supplier_No = '" & Replace(Mid(frmGRV.cmbSuppliers, InStrRev(frmGRV.cmbSuppliers, "(") + 1), ")", "") & "' and Product_Code = '" & txtProductCode.Text & "'"
        End If
        MsgBox "Product update successful", vbInformation, "HeroPOS"
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 1) = .txtDescription & " - " & txtProductCode.Text
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 7) = .txtLandCost.Text
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 8) = TaxRate & "%"
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 10) = Depart
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 3) = txtPackSize.Text
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 4) = "0"
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 5) = "0"
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 6) = "0"
        frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 9) = frmGRV.grdGRV.ValueMatrix(frmGRV.grdGRV.Row, 6) * frmGRV.grdGRV.ValueMatrix(frmGRV.grdGRV.Row, 7)
    End With
    Unload Me
End Sub
Private Sub Form_Load()
    cmbUnit.Clear
    cmbUnit.AddItem "ml"
    cmbUnit.AddItem "lt"
    cmbUnit.AddItem "g"
    cmbUnit.AddItem "kg"
    cmbUnit.AddItem "ton"
    cmbUnit.AddItem "each"
    cmbUnit.AddItem "box"
    cmbUnit.AddItem "Preparation Recipe"
    cmbUnit.Text = "each"
    cmbTax.Clear
    ActiveReadServer "Select * from Tax_Rates order by Tax_Type"
    While Not rs.EOF
        cmbTax.AddItem rs.Fields("Tax_type") & " - " & rs.Fields("Tax_Rate") & "% " & rs.Fields("Description")
        rs.MoveNext
    Wend
    rs.Close
    If cmbTax.ListCount > 1 Then
        cmbTax.Text = cmbTax.List(0)
    End If
    
    cmbDepartments.Clear
    
    ActiveReadServer "Select * from Departments order by Department_No"
    cmbDepartments.AddItem "<Unbound>"
    While Not rs.EOF
        cmbDepartments.AddItem rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        rs.MoveNext
    Wend
    rs.Close
    cmbDepartments.Text = "<Unbound>"
End Sub
Private Sub Form_Activate()
    ActiveReadServer "Select *,(Select Description from Tax_Rates where Products.Tax_Type = Tax_Rates.Tax_Type) as Tax_Description,(Select Dept_Name from Departments where Products.Department_No = Departments.Department_No) as Dept_Name from Products where Product_Code='" & frmGRV.grdGRV.TextMatrix(frmGRV.grdGRV.Row, 0) & "'"
        If rs.RecordCount > 0 Then
            frmProd.Caption = "Edit a Product..."
            txtProductCode.Text = rs.Fields("Product_Code")
            TopRow = 0
            BottomRow = 0
            For i = 1 To Len(txtProductCode.Text) - 1
                If Len(txtProductCode.Text) < 9 Then
                    If i / 2 <> Int(i / 2) Then
                        TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
                    Else
                        BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
                    End If
                Else
                    If i / 2 = Int(i / 2) Then
                        TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
                    Else
                        BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
                    End If
                End If
            Next i
            TopRow = TopRow * 3
            result = TopRow + BottomRow
            result = (1 - ((result / 10) - Int((result / 10)))) * 10
            If result = 10 Then result = 0
            If Round(result, 0) = Int(Val(Right(txtProductCode, 1))) Then
                PicBC.Visible = True
            Else
                PicBC.Visible = False
            End If
            txtDescription.Text = rs.Fields("Description")
            txtPackSize.Text = rs.Fields("Pack_Size") & ""
            txtShort.Text = rs.Fields("Short_Description")
            If rs.Fields("Unit_of_Measure") & "" = "" Then
                cmbUnit.Text = "each"
            Else
                cmbUnit.Text = rs.Fields("Unit_of_Measure") & ""
            End If
            If rs.Fields("Unit_Size") = 0 Then
                txtUnitSize.Text = ""
            Else
                txtUnitSize.Text = rs.Fields("Unit_Size") & ""
            End If
            cmbDepartments.Text = rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
            txtLandCost.Text = Format(rs.Fields("Landed_Cost"), "0.00")
            txtLandCost.ToolTipText = " Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & " "
            cmbTax.Text = rs.Fields("Tax_type") & " - " & rs.Fields("Sales_Tax") & "% " & rs.Fields("Tax_Description")
            txtSellIncl.Text = Format(rs.Fields("Selling_Price"), "0.00")
            txtSellIncl.Tag = "1"
            chkStock.Value = rs.Fields("Stock_Item")
            chkRecipe.Value = rs.Fields("Recipe_Item")
            chkSales.Value = rs.Fields("Sales_Item")
            chkTouch.Value = rs.Fields("Touch_Item")
            chkDeposit.Value = rs.Fields("Returnable_Item")
            chkScale.Value = rs.Fields("Scale_Item")
            chkDelete.Value = rs.Fields("Once_Off")
            chkWhole.Value = Val(rs.Fields("Whole_Unit") & "")
            If chkSales.Value = 1 Then
                txtSellExcl.Tag = "1"
                Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
                If txtSellIncl.Text <> "N/A" Then
                    txtSellExcl.Text = Format(txtSellIncl.Text / ((100 + Tax) / 100), "0.00")
                End If
                txtSellExcl.Tag = ""
                If Val(txtLandCost.Text) <> 0 Then
                    txtMarkup.Tag = "1"
                    txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                    txtMarkup.Tag = ""
                End If
                If Val(txtSellExcl) <> 0 Then
                    txtGross.Tag = "1"
                    txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
                    txtGross.Tag = ""
                Else
                    txtGross.Tag = "1"
                    txtGross.Text = "0"
                    txtGross.Tag = ""
                End If
                chkTouch.Enabled = True
                txtSellExcl.Enabled = True
                txtSellIncl.Enabled = True
                txtMarkup.Enabled = True
                txtGross.Enabled = True
                cmbTax.Enabled = True
                ActiveReadServer2 "Select * from Suppliers_Links_View where Supplier_No = '" & Replace(Mid(frmGRV.cmbSuppliers, InStrRev(frmGRV.cmbSuppliers, "(") + 1), ")", "") & "' and Product_Code = '" & txtProductCode.Text & "'"
                If rs2.RecordCount > 0 Then
                    txtSupp.Text = rs2.Fields("Supplier_Code")
                End If
                rs2.Close
            End If
            txtSellIncl.Tag = ""
            txtDescription.SetFocus
        Else
            frmProd.Caption = "Create a Product..."
            txtDescription.Text = ""
            txtPackSize.Text = "1"
            txtShort.Text = ""
            cmbUnit.Text = "each"
            txtUnitSize.Text = ""
            txtLandCost.Text = "0.00"
            txtMarkup.Text = "0"
            txtGross.Text = "0"
            txtSellExcl.Text = "0.00"
            txtSellIncl.Text = "0.00"
            txtProductCode.SetFocus
        End If
        rs.Close
        CheckforSave
End Sub
Private Sub cmbTax_Change()
    Calculate_Plu "Tax"
End Sub

Private Sub txtDescription_Change()
    CheckforSave
End Sub
Private Sub txtDescription_GotFocus()
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
End Sub
Private Sub txtEmpty_GotFocus()
    txtEmpty.SelStart = 0
    txtEmpty.SelLength = Len(txtEmpty.Text)
End Sub

Private Sub txtEmpty_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtGross_Change()
    If txtGross.Tag = "" Then Calculate_Plu "GP"
End Sub
Private Sub txtLandCost_Change()
    Calculate_Plu "Landed Cost"
End Sub
Private Sub txtMarkup_Change()
    If txtMarkup.Tag = "" Then Calculate_Plu "Markup"
End Sub
Private Sub txtPackSize_GotFocus()
    txtPackSize.SelStart = 0
    txtPackSize.SelLength = Len(txtPackSize.Text)
End Sub
Private Sub txtPackSize_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtProductCode_GotFocus()
    txtProductCode.SelStart = 0
    txtProductCode.SelLength = Len(txtProductCode.Text)
End Sub

Private Sub txtProductCode_LostFocus()
    ActiveReadServer "Select *,(Select Description from Tax_Rates where Products.Tax_Type = Tax_Rates.Tax_Type) as Tax_Description,(Select Dept_Name from Departments where Products.Department_No = Departments.Department_No) as Dept_Name from Products where Product_Code='" & txtProductCode.Text & "'"
    If rs.RecordCount > 0 Then
        TopRow = 0
        BottomRow = 0
        For i = 1 To Len(txtProductCode.Text) - 1
            If Len(txtProductCode.Text) < 9 Then
                If i / 2 <> Int(i / 2) Then
                    TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
                Else
                    BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
                End If
            Else
                If i / 2 = Int(i / 2) Then
                    TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
                Else
                    BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
                End If
            End If
        Next i
        TopRow = TopRow * 3
        result = TopRow + BottomRow
        result = (1 - ((result / 10) - Int((result / 10)))) * 10
        If result = 10 Then result = 0
        If Round(result, 0) = Int(Val(Right(txtProductCode, 1))) Then
            PicBC.Visible = True
        Else
            PicBC.Visible = False
        End If
        txtDescription.Text = rs.Fields("Description")
        txtPackSize.Text = rs.Fields("Pack_Size") & ""
        txtShort.Text = rs.Fields("Short_Description")
        If rs.Fields("Unit_of_Measure") & "" = "" Then
            cmbUnit.Text = "each"
        Else
            cmbUnit.Text = rs.Fields("Unit_of_Measure") & ""
        End If
        If rs.Fields("Unit_Size") = 0 Then
            txtUnitSize.Text = ""
        Else
            txtUnitSize.Text = rs.Fields("Unit_Size") & ""
        End If
        cmbDepartments.Text = rs.Fields("Department_No") & " - " & rs.Fields("Dept_Name")
        txtLandCost.Text = Format(rs.Fields("Landed_Cost"), "0.00")
        txtLandCost.ToolTipText = " Average Cost: " & Format(rs.Fields("Ave_Cost"), "0.00") & " "
        cmbTax.Text = rs.Fields("Tax_type") & " - " & rs.Fields("Sales_Tax") & "% " & rs.Fields("Tax_Description")
        txtSellIncl.Text = Format(rs.Fields("Selling_Price"), "0.00")
        txtSellIncl.Tag = "1"
        chkStock.Value = rs.Fields("Stock_Item")
        chkRecipe.Value = rs.Fields("Recipe_Item")
        chkSales.Value = rs.Fields("Sales_Item")
        chkTouch.Value = rs.Fields("Touch_Item")
        chkDeposit.Value = rs.Fields("Returnable_Item")
        chkScale.Value = rs.Fields("Scale_Item")
        chkDelete.Value = rs.Fields("Once_Off")
        chkWhole.Value = Val(rs.Fields("Whole_Unit") & "")
        If chkSales.Value = 1 Then
            txtSellExcl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If txtSellIncl.Text <> "N/A" Then
                txtSellExcl.Text = Format(txtSellIncl.Text / ((100 + Tax) / 100), "0.00")
            End If
            txtSellExcl.Tag = ""
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                txtMarkup.Tag = ""
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
                txtGross.Tag = ""
            Else
                txtGross.Tag = "1"
                txtGross.Text = "0"
                txtGross.Tag = ""
            End If
            chkTouch.Enabled = True
            txtSellExcl.Enabled = True
            txtSellIncl.Enabled = True
            txtMarkup.Enabled = True
            txtGross.Enabled = True
            cmbTax.Enabled = True
            ActiveReadServer2 "Select * from Suppliers_Links_View where Supplier_No = '" & Replace(Mid(frmGRV.cmbSuppliers, InStrRev(frmGRV.cmbSuppliers, "(") + 1), ")", "") & "' and Product_Code = '" & txtProductCode.Text & "'"
                If rs2.RecordCount > 0 Then
                    txtSupp.Text = rs2.Fields("Supplier_Code")
                End If
                rs2.Close
        End If
        txtSellIncl.Tag = ""
    Else
        frmProd.Caption = "Create a Product..."
        txtDescription.Text = ""
        txtPackSize.Text = "1"
        txtShort.Text = ""
        cmbUnit.Text = "each"
        txtUnitSize.Text = ""
        txtLandCost.Text = "0.00"
        txtMarkup.Text = "0"
        txtGross.Text = "0"
        txtSellExcl.Text = "0.00"
        txtSellIncl.Text = "0.00"
        txtDescription.SetFocus
    End If
    rs.Close
End Sub

Private Sub txtSellExcl_Change()
    If txtSellExcl.Tag = "" Then Calculate_Plu "SellExcl"
End Sub
Private Sub txtSellIncl_Change()
    If txtSellIncl.Tag = "" Then Calculate_Plu "SellIncl"
End Sub
Private Sub txtShort_Change()
    CheckforSave
End Sub
Private Sub txtShort_GotFocus()
    txtShort.SelStart = 0
    txtShort.SelLength = Len(txtShort.Text)
End Sub
Private Sub txtShort_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            txtDescription.SetFocus
        Case 40
            cmbUnit.SetFocus
    End Select
End Sub
Private Sub txtShort_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtSupp_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub
Private Sub txtSupp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case 39
            KeyAscii = 0
        Case 65 To 90
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtUnitSize_GotFocus()
    txtUnitSize.SelStart = 0
    txtUnitSize.SelLength = Len(txtUnitSize.Text)
End Sub
Private Sub txtUnitSize_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            cmbUnit.SetFocus
        Case 40
            cmbDepartments.SetFocus
    End Select
End Sub
Private Sub txtUnitSize_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Public Sub Calculate_Plu(FromWhere)
    On Error Resume Next
    If txtSellIncl.Enabled = False Then Exit Sub
    Select Case FromWhere
        Case "Landed Cost"
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                txtMarkup.Tag = ""
            Else
                 txtMarkup.Text = "0"
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
                txtGross.Tag = ""
            End If
        Case "Markup"
            If Val(txtLandCost.Text) <> 0 Then
                txtSellExcl.Tag = "1"
                txtSellExcl.Text = Format(Val(txtLandCost.Text) * ((100 + Val(txtMarkup.Text)) / 100), "0.00")
                txtSellExcl.Tag = ""
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl) * 100), 3)
                txtGross.Tag = ""
            Else
                txtGross.Tag = "1"
                txtGross.Text = "0"
                txtGross.Tag = ""
            End If
            txtSellIncl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If Val(txtLandCost.Text) <> 0 Then
                txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
            End If
            txtSellIncl.Tag = ""
        Case "GP"
            txtSellExcl.Tag = "1"
            txtSellExcl.Text = Format(txtLandCost.Text / ((100 - Val(txtGross.Text)) / 100), "0.00")
            txtSellExcl.Tag = ""
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                txtMarkup.Tag = ""
            End If
            txtSellIncl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
            txtSellIncl.Tag = ""
        Case "SellExcl"
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost) * 100), 3)
                txtMarkup.Tag = ""
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl) * 100), 3)
                txtGross.Tag = ""
            Else
                txtGross.Tag = "1"
                txtGross.Text = "0"
                txtGross.Tag = ""
            End If
            txtSellIncl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If txtSellExcl.Text <> "N/A" Then
                txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
            End If
            txtSellIncl.Tag = ""
        Case "Tax"
            txtSellIncl.Tag = "1"
            If cmbTax.Text <> "" Then
                Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            End If
            txtSellIncl.Text = Format(txtSellExcl.Text * ((100 + Tax) / 100), "0.00")
            txtSellIncl.Tag = ""
        Case "SellIncl"
            txtSellExcl.Tag = "1"
            Tax = Mid(cmbTax.Text, InStr(cmbTax.Text, "-") + 2, InStr(cmbTax.Text, "%") - InStr(cmbTax.Text, "-") - 2)
            If txtSellIncl.Text <> "N/A" Then
                txtSellExcl.Text = Format(txtSellIncl.Text / ((100 + Tax) / 100), "0.00")
            End If
            txtSellExcl.Tag = ""
            If Val(txtLandCost.Text) <> 0 Then
                txtMarkup.Tag = "1"
                txtMarkup.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtLandCost.Text) * 100), 3)
                txtMarkup.Tag = ""
            End If
            If Val(txtSellExcl) <> 0 Then
                txtGross.Tag = "1"
                txtGross.Text = Round(((Val(txtSellExcl.Text) - Val(txtLandCost.Text)) / Val(txtSellExcl.Text) * 100), 3)
                txtGross.Tag = ""
            Else
                txtGross.Tag = "1"
                txtGross.Text = "0"
                txtGross.Tag = ""
            End If
    End Select
    On Error GoTo 0
End Sub
Private Sub CheckforSave()
    If txtProductCode.Text <> "" And txtDescription.Text <> "" Then
        cmdSave.Enabled = True
        cmdBarcode.Enabled = True
    Else
        cmdSave.Enabled = False
        cmdBarcode.Enabled = False
    End If
End Sub
Public Sub chkRecipe_Click()
    Select Case chkRecipe.Value
        Case 0
            txtLandCost.Locked = False
            txtLandCost.Enabled = True
            lblCost.Caption = "Landed Cost:"
        Case 1
            Select Case cmbUnit.Text
                Case "Preparation Recipe"
                    txtLandCost.Enabled = False
                    txtLandCost.Text = "N/A"
                Case Else
                    txtLandCost.Locked = True
                    lblCost.Caption = "Theoretical Cost:"
            End Select
    End Select
End Sub

Public Sub chkSales_Click()
    Select Case chkSales.Value
        Case 0
            chkTouch.Enabled = False
            chkTouch.Value = 0
            txtSellExcl.Text = "N/A"
            txtSellIncl.Text = "N/A"
            txtMarkup.Text = "N/A"
            txtGross.Text = "N/A"
            txtMarkup.Enabled = False
            txtGross.Enabled = False
            txtSellExcl.Enabled = False
            txtSellIncl.Enabled = False
            cmbTax.Enabled = False
        Case 1
            If txtSellIncl.Tag = "1" Then Exit Sub
            chkTouch.Enabled = True
            txtSellExcl.Enabled = True
            txtSellIncl.Enabled = True
            txtMarkup.Enabled = True
            txtGross.Enabled = True
            txtMarkup.Text = "0"
            txtGross.Text = "0"
            txtSellExcl.Text = "0.00"
            txtSellIncl.Text = "0.00"
            cmbTax.Enabled = True
    End Select
End Sub
Public Sub chkStock_Click()
    Select Case chkStock.Value
        Case 0
            chkRecipe.Enabled = True
            chkDelete.Enabled = False
            chkDelete.Value = 0
            chkDeposit.Value = 0
        Case 1
            chkRecipe.Enabled = False
            chkRecipe.Value = 0
            chkDelete.Enabled = True
    End Select
End Sub
Private Sub cmbDepartments_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbDepartments_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
                If txtProductCode.Text = "" Then txtProductCode.SetFocus
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                If txtUnitSize.Enabled = True Then
                    txtUnitSize.SetFocus
                Else
                    cmbUnit.SetFocus
                End If
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
End Sub

Private Sub cmbTax_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbTax_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtSellExcl.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
End Sub
Private Sub cmbUnit_Change()
    If cmbUnit.Text = "each" Or cmbUnit.Text = "Preparation Recipe" Then
        txtUnitSize.Text = ""
        txtUnitSize.Enabled = False
    Else
        txtUnitSize.Enabled = True
        chkStock.Enabled = True
        chkSales.Enabled = True
    End If
    If cmbUnit.Text = "Preparation Recipe" Then
        chkStock.Value = 0
        chkSales.Value = 0
        chkStock.Enabled = False
        chkSales.Enabled = False
    Else
        chkStock.Enabled = True
        chkSales.Enabled = True
    End If
End Sub
Private Sub cmbUnit_DropButtonClick()
    Select Case ActiveControl.Tag
        Case "Dropped"
            ActiveControl.Tag = "Up"
        Case "Up"
            ActiveControl.Tag = "Dropped"
    End Select
End Sub
Private Sub cmbUnit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case 13
            If ActiveControl.Tag = "Up" Then
                ActiveControl.DropDown
                KeyCode = 0
            End If
        Case 38
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                txtShort.SetFocus
            Else
                If ActiveControl.ListIndex = 0 Then KeyCode = 0
            End If
        Case 40
            If ActiveControl.Tag = "Up" Then
                KeyCode = 0
                SendKeys "{TAB}"
            Else
                If ActiveControl.ListIndex = ActiveControl.ListCount - 1 Then KeyCode = 0
            End If
    End Select
End Sub
Private Sub cmdBarcode_Click()
    On Error Resume Next
    frmGRV.Tag = "1"
    Load frmProdChange
    frmProdChange.Tag = "GRV"
    frmProdChange.Show vbModal
    On Error Resume Next
End Sub
Private Sub txtProductCode_Change()
    BottomRow = 0
    For i = 1 To Len(txtProductCode.Text) - 1
        If Len(txtProductCode.Text) < 9 Then
            If i / 2 <> Int(i / 2) Then
                TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
            End If
        Else
            If i / 2 = Int(i / 2) Then
                TopRow = TopRow + Val(Mid(txtProductCode.Text, i, 1))
            Else
                BottomRow = BottomRow + Val(Mid(txtProductCode.Text, i, 1))
            End If
        End If
    Next i
    TopRow = TopRow * 3
    result = TopRow + BottomRow
    result = (1 - ((result / 10) - Int((result / 10)))) * 10
    If result = 10 Then result = 0
    If Round(result, 0) = Int(Val(Right(txtProductCode, 1))) Then
        PicBC.Visible = True
    Else
        PicBC.Visible = False
    End If
    CheckforSave
    End Sub
Private Sub txtproductcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            If txtSellIncl.Enabled = True Then
                txtSellIncl.SetFocus
            Else
                txtLandCost.SetFocus
            End If
        Case 40
            txtDescription.SetFocus
    End Select
End Sub
Private Sub txtproductcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case 39
            KeyAscii = 0
        Case 65 To 90
        Case 97 To 122
            KeyAscii = KeyAscii - 32
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub txtSellExcl_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtSellExcl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            txtGross.SetFocus
        Case 40
            cmbTax.SetFocus
    End Select
End Sub

Private Sub txtSellExcl_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtSellExcl_LostFocus()
    If txtSellExcl.Text = "" Then txtSellExcl.Text = "0.00"
    txtSellExcl.Text = Format(txtSellExcl.Text, "0.00")
End Sub

Private Sub txtSellIncl_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtSellIncl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            cmbTax.SetFocus
        Case 40
            txtProductCode.SetFocus
    End Select
End Sub

Private Sub txtSellIncl_KeyPress(KeyAscii As Integer)
    If InStr(txtSellIncl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtSellIncl_LostFocus()
    If txtSellIncl.Text = "" Then txtSellIncl.Text = "0.00"
    txtSellIncl.Text = Format(txtSellIncl.Text, "0.00")
End Sub
Private Sub txtMarkup_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub
Private Sub txtMarkup_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            txtLandCost.SetFocus
        Case 40
            txtGross.SetFocus
    End Select
End Sub
Private Sub txtMarkup_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtMarkup_LostFocus()
    If txtMarkup.Text = "" Then txtMarkup.Text = "0"
End Sub
Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            txtProductCode.SetFocus
        Case 40
            txtShort.SetFocus
    End Select
End Sub
Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub txtDescription_LostFocus()
    On Error Resume Next
    CheckforSave
    txtDescription.Text = UCase(Left(txtDescription.Text, 1)) & Mid(txtDescription.Text, 2)
    If Len(txtDescription.Text) < 26 And txtShort.Text = "" Then
        txtShort.Text = txtDescription.Text
    End If
    On Error GoTo 0
End Sub

Private Sub txtGross_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtGross_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            txtMarkup.SetFocus
        Case 40
            txtSellExcl.SetFocus
    End Select
End Sub

Private Sub txtGross_KeyPress(KeyAscii As Integer)
    If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtGross_LostFocus()
    If txtGross.Text = "" Then txtGross.Text = "0"
End Sub
Private Sub txtLandCost_GotFocus()
    txtLandCost.SelStart = 0
    txtLandCost.SelLength = Len(txtLandCost.Text)
End Sub
Private Sub txtLandCost_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 38
            cmbDepartments.SetFocus
        Case 40
            If txtMarkup.Enabled = True Then
                txtMarkup.SetFocus
            Else
                txtProductCode.SetFocus
            End If
    End Select
End Sub

Private Sub txtLandCost_KeyPress(KeyAscii As Integer)
    If InStr(txtLandCost.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtLandCost_LostFocus()
    If txtLandCost.Text = "" Then txtLandCost.Text = "0.00"
    txtLandCost.Text = Format(txtLandCost.Text, "0.00")
    Calculate_Plu "Landed Cost"
End Sub



