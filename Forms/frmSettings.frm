VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Settings"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6075
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10716
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Connectivity"
      TabPicture(0)   =   "frmSettings.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtServer"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image3(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "opbAuth(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "opbAuth(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbDatabase"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtLogin"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPassword"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtErrorLog"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtUserLog"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label6"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdForms(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdErrorLog(4)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdUserLog(3)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Workstation Settings"
      TabPicture(1)   =   "frmSettings.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtWNumber"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtWork_Name"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmbPrinter"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4(4)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label4(5)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmbKitchen"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label4(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmbLoc"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkSOH"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmbPort"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkCon"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkPrintZero"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chkVoids"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chkTransfers"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkReplicate"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkAskLoc"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "chkBarStock"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Frame2"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Frame8"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "General"
      TabPicture(2)   =   "frmSettings.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtVat"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtSwiss"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame1(0)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame7"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame9"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame10"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame1(1)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Stock && Sales"
      TabPicture(3)   =   "frmSettings.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdStock"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Devices"
      TabPicture(4)   =   "frmSettings.frx":007C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Image1(3)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame3"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame5"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame11"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Frame6"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Conversion Message"
         ForeColor       =   &H80000008&
         Height          =   1305
         Index           =   1
         Left            =   -68760
         TabIndex        =   120
         Top             =   3990
         Width           =   3015
         Begin VB.Label Label18 
            BackColor       =   &H80000005&
            Caption         =   "*   0 = disabled"
            Height          =   255
            Left            =   360
            TabIndex        =   126
            Top             =   920
            Width           =   1695
         End
         Begin MSForms.TextBox txtCurRate 
            Height          =   315
            Left            =   1680
            TabIndex        =   125
            Top             =   480
            Width           =   735
            VariousPropertyBits=   746604571
            MaxLength       =   20
            BorderStyle     =   1
            Size            =   "1296;556"
            Value           =   "0.00"
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtCurr 
            Height          =   315
            Left            =   480
            TabIndex        =   124
            Top             =   480
            Width           =   855
            VariousPropertyBits=   746604571
            MaxLength       =   20
            BorderStyle     =   1
            Size            =   "1508;556"
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ZAR"
            Height          =   315
            Left            =   2520
            TabIndex        =   123
            Top             =   540
            Width           =   375
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "One"
            Height          =   315
            Left            =   120
            TabIndex        =   122
            Top             =   540
            Width           =   280
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "= "
            Height          =   315
            Left            =   1440
            TabIndex        =   121
            Top             =   540
            Width           =   135
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label Printer"
         ForeColor       =   &H00000000&
         Height          =   2235
         Left            =   5820
         TabIndex        =   107
         Top             =   540
         Width           =   3555
         Begin btButtonEx.ButtonEx cmdTestLabel 
            Height          =   315
            Left            =   2190
            TabIndex        =   108
            Top             =   1860
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            Appearance      =   3
            AutoMask        =   0   'False
            Caption         =   "Test"
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
         Begin MSForms.ComboBox cmbLabel 
            Height          =   315
            Left            =   1230
            TabIndex        =   116
            Top             =   180
            Width           =   2175
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3836;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   345
            Index           =   15
            Left            =   60
            TabIndex        =   115
            Top             =   240
            Width           =   1125
            BackColor       =   16777215
            Caption         =   "Label Printer:"
            Size            =   "1984;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.Label Label11 
            Height          =   285
            Left            =   240
            TabIndex        =   114
            Top             =   840
            Width           =   1215
            BackColor       =   16777215
            Caption         =   "Barcode Height:"
            Size            =   "2143;503"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtLabel 
            Height          =   315
            Left            =   1530
            TabIndex        =   113
            Top             =   810
            Width           =   1905
            VariousPropertyBits=   746604571
            MaxLength       =   20
            BorderStyle     =   1
            Size            =   "3360;556"
            Value           =   "0"
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtwidth 
            Height          =   315
            Left            =   1530
            TabIndex        =   112
            Top             =   1260
            Width           =   1905
            VariousPropertyBits=   746604571
            MaxLength       =   20
            BorderStyle     =   1
            Size            =   "3360;556"
            Value           =   "0"
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label12 
            Height          =   285
            Left            =   240
            TabIndex        =   111
            Top             =   1290
            Width           =   1215
            BackColor       =   16777215
            Caption         =   "Label Width:"
            Size            =   "2143;503"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Small Price Label Properties:"
            Height          =   225
            Left            =   240
            TabIndex        =   110
            Top             =   570
            Width           =   2325
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Shelf talkers have a set size"
            Height          =   375
            Left            =   240
            TabIndex        =   109
            Top             =   1800
            Width           =   1125
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Red Printing Department"
         ForeColor       =   &H00000000&
         Height          =   1845
         Left            =   5850
         TabIndex        =   104
         Top             =   3420
         Width           =   3555
         Begin VB.CheckBox ChkRkitchen 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Print prices on Kitchen messages:"
            Height          =   315
            Left            =   180
            TabIndex        =   119
            Top             =   1380
            Width           =   3225
         End
         Begin MSForms.ComboBox CmbRed2 
            Height          =   315
            Left            =   1530
            TabIndex        =   118
            Top             =   870
            Width           =   1905
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3360;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   195
            Index           =   16
            Left            =   60
            TabIndex        =   117
            Top             =   960
            Width           =   1395
            BackColor       =   16777215
            Caption         =   "Sub Department1:"
            Size            =   "2461;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox CmbRed 
            Height          =   315
            Left            =   1530
            TabIndex        =   106
            Top             =   390
            Width           =   1905
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3360;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   195
            Index           =   10
            Left            =   60
            TabIndex        =   105
            Top             =   480
            Width           =   1395
            BackColor       =   16777215
            Caption         =   "Sub Department1:"
            Size            =   "2461;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Options"
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   -69750
         TabIndex        =   94
         Top             =   720
         Width           =   4065
         Begin MSForms.CheckBox chkZero 
            Height          =   405
            Left            =   180
            TabIndex        =   103
            Top             =   2730
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "Do not add Zero Sales to Statements."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkMember 
            Height          =   405
            Left            =   180
            TabIndex        =   102
            Top             =   2415
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "Ask Member Number after each Sale."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkTradePrint 
            Height          =   405
            Left            =   180
            TabIndex        =   101
            Top             =   2100
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "Print Trade Analysis on Slip Printer."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkChargePrint 
            Height          =   405
            Left            =   180
            TabIndex        =   100
            Top             =   1785
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "Print Two Slips when Charging to Debtors."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkQCash 
            Height          =   405
            Left            =   180
            TabIndex        =   99
            Top             =   1470
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "Use Quick Cashup Method."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkBarcode 
            Height          =   405
            Left            =   180
            TabIndex        =   98
            Top             =   1155
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "Print Barcode on Stock Count Sheets."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkVoidReason 
            Height          =   405
            Left            =   180
            TabIndex        =   97
            Top             =   840
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "Supply Reasons for Voids."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkService 
            Height          =   405
            Left            =   180
            TabIndex        =   96
            Top             =   525
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "Use Service Charges when Tendering."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkAccess 
            Height          =   405
            Left            =   180
            TabIndex        =   95
            Top             =   210
            Width           =   3525
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6218;714"
            Value           =   "0"
            Caption         =   "Only use User Password for System Access."
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Print with Trade Analysis"
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   -73290
         TabIndex        =   90
         Top             =   2430
         Width           =   3405
         Begin MSForms.CheckBox chkPayout 
            Height          =   405
            Left            =   180
            TabIndex        =   93
            Top             =   930
            Width           =   3135
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "5530;714"
            Value           =   "0"
            Caption         =   "Payout Summary"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkRA 
            Height          =   405
            Left            =   180
            TabIndex        =   92
            Top             =   600
            Width           =   3135
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "5530;714"
            Value           =   "0"
            Caption         =   "Receive on Account Summary"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkCharge 
            Height          =   405
            Left            =   180
            TabIndex        =   91
            Top             =   270
            Width           =   3135
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "5530;714"
            Value           =   "0"
            Caption         =   "Debtor Sales Summary"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Active Discount Buttons"
         ForeColor       =   &H00000000&
         Height          =   3435
         Left            =   -69210
         TabIndex        =   76
         Top             =   2100
         Width           =   2265
         Begin MSForms.CheckBox chkDiscount 
            Height          =   345
            Index           =   9
            Left            =   210
            TabIndex        =   86
            Top             =   2895
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;609"
            Value           =   "0"
            Caption         =   "Free"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   8
            Left            =   210
            TabIndex        =   85
            Top             =   2640
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Cost +10"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   7
            Left            =   210
            TabIndex        =   84
            Top             =   2340
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Less 80%"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   6
            Left            =   210
            TabIndex        =   83
            Top             =   2040
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Less 70%"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   5
            Left            =   210
            TabIndex        =   82
            Top             =   1740
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Less 60%"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   4
            Left            =   210
            TabIndex        =   81
            Top             =   1440
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Less 50%"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   3
            Left            =   210
            TabIndex        =   80
            Top             =   1140
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Less 40%"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   2
            Left            =   210
            TabIndex        =   79
            Top             =   840
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Less 30%"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   1
            Left            =   210
            TabIndex        =   78
            Top             =   540
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Less 20%"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkDiscount 
            Height          =   300
            Index           =   0
            Left            =   210
            TabIndex        =   77
            Top             =   240
            Width           =   2265
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3995;529"
            Value           =   "0"
            Caption         =   "Less 10%"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cost Matrix"
         ForeColor       =   &H80000008&
         Height          =   1305
         Left            =   -73290
         TabIndex        =   71
         Top             =   3990
         Width           =   4425
         Begin VSFlex8Ctl.VSFlexGrid grdCost 
            Height          =   690
            Left            =   180
            TabIndex        =   72
            Top             =   420
            Width           =   4065
            _cx             =   7170
            _cy             =   1217
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
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
            BackColorSel    =   16506073
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   330
            RowHeightMax    =   0
            ColWidthMin     =   400
            ColWidthMax     =   400
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSettings.frx":0098
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
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
            ExplorerBar     =   0
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
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Customer Display"
         ForeColor       =   &H00000000&
         Height          =   1665
         Left            =   240
         TabIndex        =   51
         Top             =   4050
         Width           =   4215
         Begin MSForms.ComboBox cmbDisplaySet 
            Height          =   315
            Left            =   1800
            TabIndex        =   68
            Top             =   1140
            Width           =   2265
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   345
            Index           =   14
            Left            =   30
            TabIndex        =   67
            Top             =   1200
            Width           =   1725
            BackColor       =   16777215
            Caption         =   "Settings:"
            Size            =   "3043;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox cmbDisplayPort 
            Height          =   315
            Left            =   1800
            TabIndex        =   66
            Top             =   720
            Width           =   2265
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   345
            Index           =   13
            Left            =   30
            TabIndex        =   65
            Top             =   780
            Width           =   1725
            BackColor       =   16777215
            Caption         =   "Comms Port:"
            Size            =   "3043;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox cmbDisplay 
            Height          =   315
            Left            =   1800
            TabIndex        =   64
            Top             =   300
            Width           =   2265
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   345
            Index           =   12
            Left            =   30
            TabIndex        =   63
            Top             =   360
            Width           =   1725
            BackColor       =   16777215
            Caption         =   "Display Model:"
            Size            =   "3043;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Checkout Scale"
         ForeColor       =   &H00000000&
         Height          =   1665
         Left            =   240
         TabIndex        =   50
         Top             =   2295
         Width           =   4215
         Begin MSForms.ComboBox cmbScaleSet 
            Height          =   315
            Left            =   1800
            TabIndex        =   62
            Top             =   1140
            Width           =   2265
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   345
            Index           =   11
            Left            =   30
            TabIndex        =   61
            Top             =   1200
            Width           =   1725
            BackColor       =   16777215
            Caption         =   "Settings:"
            Size            =   "3043;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox cmbScalePort 
            Height          =   315
            Left            =   1800
            TabIndex        =   60
            Top             =   720
            Width           =   2265
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Texter 
            Height          =   345
            Index           =   10
            Left            =   30
            TabIndex        =   59
            Top             =   780
            Width           =   1725
            BackColor       =   16777215
            Caption         =   "Comms Port:"
            Size            =   "3043;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox cmbScale 
            Height          =   315
            Left            =   1800
            TabIndex        =   58
            Top             =   300
            Width           =   2265
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   345
            Index           =   9
            Left            =   30
            TabIndex        =   57
            Top             =   360
            Width           =   1725
            BackColor       =   16777215
            Caption         =   "Checkout Scale Model:"
            Size            =   "3043;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Drawers"
         ForeColor       =   &H00000000&
         Height          =   1605
         Left            =   240
         TabIndex        =   49
         Top             =   570
         Width           =   5505
         Begin btButtonEx.ButtonEx cmdTest 
            Height          =   315
            Index           =   0
            Left            =   4140
            TabIndex        =   69
            Top             =   360
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            Appearance      =   3
            AutoMask        =   0   'False
            Caption         =   "Test 1"
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
         Begin btButtonEx.ButtonEx cmdTest 
            Height          =   315
            Index           =   1
            Left            =   4140
            TabIndex        =   70
            Top             =   780
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            Appearance      =   3
            AutoMask        =   0   'False
            Caption         =   "Test 2"
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
         Begin MSForms.CheckBox chkDrawer 
            Height          =   405
            Left            =   1785
            TabIndex        =   56
            Top             =   1140
            Width           =   3405
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6006;714"
            Value           =   "0"
            Caption         =   "This Workstation Uses Both"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbDrawer2 
            Height          =   315
            Left            =   1800
            TabIndex        =   55
            Top             =   780
            Width           =   2265
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   345
            Index           =   8
            Left            =   30
            TabIndex        =   54
            Top             =   840
            Width           =   1725
            BackColor       =   16777215
            Caption         =   "Drawer Two Kickstring:"
            Size            =   "3043;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox cmbDrawer1 
            Height          =   315
            Left            =   1800
            TabIndex        =   53
            Top             =   360
            Width           =   2265
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   0
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   345
            Index           =   7
            Left            =   30
            TabIndex        =   52
            Top             =   420
            Width           =   1725
            BackColor       =   16777215
            Caption         =   "Drawer One Kickstring:"
            Size            =   "3043;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Slip Printer Mode"
         ForeColor       =   &H00000000&
         Height          =   1485
         Left            =   -72810
         TabIndex        =   31
         Top             =   2100
         Width           =   3525
         Begin MSForms.OptionButton opbSlip 
            Height          =   435
            Index           =   2
            Left            =   270
            TabIndex        =   44
            Top             =   930
            Width           =   3075
            BackColor       =   16777215
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "5424;767"
            Value           =   "0"
            Caption         =   "Epson Emulation Discontinued Models"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton opbSlip 
            Height          =   435
            Index           =   0
            Left            =   270
            TabIndex        =   33
            Top             =   300
            Width           =   2115
            BackColor       =   16777215
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "3731;767"
            Value           =   "1"
            Caption         =   "Epson Emulation"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton opbSlip 
            Height          =   435
            Index           =   1
            Left            =   270
            TabIndex        =   32
            Top             =   600
            Width           =   2925
            BackColor       =   16777215
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "5159;767"
            Value           =   "0"
            Caption         =   "Epson Emulation Older Models"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Department order on Touch"
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   0
         Left            =   -73290
         TabIndex        =   28
         Top             =   1260
         Width           =   3405
         Begin MSForms.OptionButton opbDept 
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   37
            Top             =   300
            Width           =   2505
            BackColor       =   16777215
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "4419;556"
            Value           =   "1"
            Caption         =   "Order by Description"
            SpecialEffect   =   0
            GroupName       =   "1"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton opbDept 
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   36
            Top             =   630
            Width           =   2505
            BackColor       =   16777215
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "4419;556"
            Value           =   "0"
            Caption         =   "Order by Number"
            SpecialEffect   =   0
            GroupName       =   "1"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin btButtonEx.ButtonEx cmdUserLog 
         Height          =   315
         Index           =   3
         Left            =   -67770
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4860
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "..."
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
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdErrorLog 
         Height          =   315
         Index           =   4
         Left            =   -67770
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   5220
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "..."
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
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdForms 
         Height          =   315
         Index           =   3
         Left            =   -70110
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2970
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "&Test Connection"
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
         ShowFocus       =   0
      End
      Begin VSFlex8Ctl.VSFlexGrid grdStock 
         Height          =   5460
         Left            =   -74910
         TabIndex        =   42
         Top             =   390
         Width           =   9435
         _cx             =   16642
         _cy             =   9631
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         BackColorSel    =   16506073
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   11
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSettings.frx":0110
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
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
      End
      Begin MSForms.CheckBox chkBarStock 
         Height          =   405
         Left            =   -72840
         TabIndex        =   89
         Top             =   5490
         Width           =   3405
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6006;714"
         Value           =   "0"
         Caption         =   "Print Barcodes on A4 Stock Count Sheets"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkAskLoc 
         Height          =   405
         Left            =   -69570
         TabIndex        =   88
         Top             =   1290
         Width           =   3405
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6006;714"
         Value           =   "0"
         Caption         =   "Select Location when Selling"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkReplicate 
         Height          =   405
         Left            =   -70500
         TabIndex        =   87
         Top             =   510
         Width           =   3405
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6006;714"
         Value           =   "0"
         Caption         =   "This workstation is a Replication Client"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkTransfers 
         Height          =   405
         Left            =   -72840
         TabIndex        =   75
         Top             =   5190
         Width           =   3405
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6006;714"
         Value           =   "0"
         Caption         =   "Print Transfers on Slip Printer"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSwiss 
         Height          =   315
         Left            =   -73290
         TabIndex        =   74
         Top             =   5370
         Width           =   2175
         VariousPropertyBits=   746604571
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "3836;556"
         Value           =   "0.05"
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label10 
         Height          =   285
         Left            =   -74790
         TabIndex        =   73
         Top             =   5400
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Swiss Rounding:"
         Size            =   "2566;503"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Image Image1 
         Height          =   5475
         Index           =   3
         Left            =   90
         Top             =   450
         Width           =   9435
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "16642;9657"
      End
      Begin MSForms.CheckBox chkVoids 
         Height          =   405
         Left            =   -72840
         TabIndex        =   48
         Top             =   4890
         Width           =   3405
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6006;714"
         Value           =   "0"
         Caption         =   "Print Voids on Slip"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkPrintZero 
         Height          =   405
         Left            =   -72840
         TabIndex        =   47
         Top             =   4575
         Width           =   3405
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6006;714"
         Value           =   "0"
         Caption         =   "Print Zero Priced Items on Slip"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkCon 
         Height          =   405
         Left            =   -72840
         TabIndex        =   46
         Top             =   4275
         Width           =   3405
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6006;714"
         Value           =   "0"
         Caption         =   "Consolidate Kitchen Printing"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbPort 
         Height          =   315
         Left            =   -69660
         TabIndex        =   45
         ToolTipText     =   " Set when Using USB to Serial Adapter "
         Top             =   1710
         Width           =   1575
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2778;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkSOH 
         Height          =   405
         Left            =   -72840
         TabIndex        =   43
         Top             =   3960
         Width           =   3405
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "6006;714"
         Value           =   "0"
         Caption         =   "Display Stock on Hand on Touch Buttons"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbLoc 
         Height          =   315
         Left            =   -72810
         TabIndex        =   41
         Top             =   1320
         Width           =   3105
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5477;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   315
         Index           =   6
         Left            =   -74730
         TabIndex        =   40
         Top             =   1350
         Width           =   1875
         BackColor       =   16777215
         Caption         =   "Situated in:"
         Size            =   "3307;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label9 
         Height          =   285
         Left            =   -74520
         TabIndex        =   39
         Top             =   840
         Width           =   1185
         BackColor       =   16777215
         Caption         =   "Vat Number:"
         Size            =   "2090;503"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtVat 
         Height          =   315
         Left            =   -73290
         TabIndex        =   38
         Top             =   810
         Width           =   3405
         VariousPropertyBits=   746604571
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "6006;556"
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbKitchen 
         Height          =   315
         Left            =   -72810
         TabIndex        =   35
         Top             =   3660
         Width           =   2535
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4471;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   315
         Index           =   5
         Left            =   -74730
         TabIndex        =   34
         Top             =   3690
         Width           =   1875
         BackColor       =   16777215
         Caption         =   "This Workstation Uses:"
         Size            =   "3307;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label4 
         Height          =   315
         Index           =   4
         Left            =   -74730
         TabIndex        =   30
         Top             =   1740
         Width           =   1875
         BackColor       =   16777215
         Caption         =   "Invoice Printer:"
         Size            =   "3307;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.ComboBox cmbPrinter 
         Height          =   315
         Left            =   -72810
         TabIndex        =   29
         Top             =   1710
         Width           =   3105
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5477;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtWork_Name 
         Height          =   315
         Left            =   -72810
         TabIndex        =   27
         Top             =   930
         Width           =   4725
         VariousPropertyBits=   746604571
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "8334;556"
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label8 
         Height          =   315
         Left            =   -74730
         TabIndex        =   26
         Top             =   960
         Width           =   1875
         BackColor       =   16777215
         Caption         =   "Workstation Name:"
         Size            =   "3307;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label7 
         Height          =   315
         Left            =   -74730
         TabIndex        =   25
         Top             =   570
         Width           =   1875
         BackColor       =   16777215
         Caption         =   "Workstation Number:"
         Size            =   "3307;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtWNumber 
         Height          =   315
         Left            =   -72810
         TabIndex        =   24
         Top             =   540
         Width           =   2175
         VariousPropertyBits=   746604571
         MaxLength       =   4
         BorderStyle     =   1
         Size            =   "3836;556"
         BorderColor     =   0
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   315
         Left            =   -74280
         TabIndex        =   23
         Top             =   4920
         Width           =   1395
         BackColor       =   16777215
         Caption         =   "User Log:"
         Size            =   "2461;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label5 
         Height          =   315
         Left            =   -74340
         TabIndex        =   22
         Top             =   5280
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Error Log:"
         Size            =   "2566;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtUserLog 
         Height          =   315
         Left            =   -72840
         TabIndex        =   6
         Top             =   4860
         Width           =   5055
         VariousPropertyBits=   746604575
         Size            =   "8916;556"
         BorderColor     =   16761024
         SpecialEffect   =   3
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtErrorLog 
         Height          =   315
         Left            =   -72840
         TabIndex        =   7
         Top             =   5220
         Width           =   5055
         VariousPropertyBits=   746604575
         Size            =   "8916;556"
         BorderColor     =   16761024
         SpecialEffect   =   3
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line3 
         X1              =   -73650
         X2              =   -68910
         Y1              =   4650
         Y2              =   4650
      End
      Begin VB.Line Line1 
         X1              =   -73650
         X2              =   -68910
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line2 
         X1              =   -73650
         X2              =   -68910
         Y1              =   1560
         Y2              =   1560
      End
      Begin MSForms.Label Label4 
         Height          =   315
         Index           =   3
         Left            =   -74400
         TabIndex        =   21
         Top             =   4590
         Width           =   885
         BackColor       =   16777215
         Caption         =   "Log Files"
         Size            =   "1561;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPassword 
         Height          =   315
         Left            =   -72840
         TabIndex        =   4
         Top             =   2970
         Width           =   2655
         VariousPropertyBits=   746604571
         Size            =   "4683;556"
         PasswordChar    =   42
         BorderColor     =   16761024
         SpecialEffect   =   3
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtLogin 
         Height          =   315
         Left            =   -72840
         TabIndex        =   3
         Top             =   2610
         Width           =   2655
         VariousPropertyBits=   746604571
         Size            =   "4683;556"
         BorderColor     =   16761024
         SpecialEffect   =   3
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   315
         Index           =   2
         Left            =   -74550
         TabIndex        =   19
         Top             =   1500
         Width           =   885
         BackColor       =   16777215
         Caption         =   "Connection"
         Size            =   "1561;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Left            =   -73800
         TabIndex        =   20
         Top             =   900
         Width           =   735
         BackColor       =   16777215
         Caption         =   "Server:"
         Size            =   "1296;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label4 
         Height          =   315
         Index           =   1
         Left            =   -74310
         TabIndex        =   18
         Top             =   3540
         Width           =   885
         BackColor       =   16777215
         Caption         =   "Options"
         Size            =   "1561;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbDatabase 
         Height          =   315
         Left            =   -73020
         TabIndex        =   5
         Top             =   3960
         Width           =   2775
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "4895;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   3
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   315
         Index           =   0
         Left            =   -73830
         TabIndex        =   17
         Top             =   3990
         Width           =   1605
         BackColor       =   16777215
         Caption         =   "Database:"
         Size            =   "2831;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   315
         Left            =   -74340
         TabIndex        =   16
         Top             =   3030
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Password:"
         Size            =   "2566;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   315
         Left            =   -74280
         TabIndex        =   15
         Top             =   2670
         Width           =   1395
         BackColor       =   16777215
         Caption         =   "Login Name:"
         Size            =   "2461;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.OptionButton opbAuth 
         Height          =   315
         Index           =   1
         Left            =   -74160
         TabIndex        =   2
         Top             =   2130
         Width           =   2505
         BackColor       =   16777215
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "4419;556"
         Value           =   "0"
         Caption         =   "Use Server authentication"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton opbAuth 
         Height          =   315
         Index           =   0
         Left            =   -74160
         TabIndex        =   1
         Top             =   1800
         Width           =   2505
         BackColor       =   16777215
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "4419;556"
         Value           =   "0"
         Caption         =   "Use Windows authentication"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image Image3 
         Height          =   765
         Index           =   0
         Left            =   -74430
         Top             =   780
         Width           =   735
         BackColor       =   16777215
         BorderStyle     =   0
         Size            =   "1296;1349"
      End
      Begin MSForms.TextBox txtServer 
         Height          =   315
         Left            =   -73020
         TabIndex        =   0
         Top             =   870
         Width           =   2655
         VariousPropertyBits=   746604571
         Size            =   "4683;556"
         BorderColor     =   16761024
         SpecialEffect   =   3
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image Image1 
         Height          =   5475
         Index           =   2
         Left            =   -74910
         Top             =   420
         Width           =   9435
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "16642;9657"
      End
      Begin MSForms.Image Image1 
         Height          =   5565
         Index           =   1
         Left            =   -74910
         Top             =   420
         Width           =   9435
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "16642;9816"
      End
      Begin MSForms.Image Image1 
         Height          =   5475
         Index           =   0
         Left            =   -74910
         Top             =   420
         Width           =   9435
         BackColor       =   16777215
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "16642;9657"
      End
   End
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   0
      Left            =   6870
      TabIndex        =   9
      Top             =   6240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   1
      Left            =   8280
      TabIndex        =   10
      Top             =   6240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Appearance      =   3
      AutoMask        =   0   'False
      Caption         =   "&Help"
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
   Begin btButtonEx.ButtonEx cmdForms 
      Height          =   345
      Index           =   2
      Left            =   5490
      TabIndex        =   8
      Top             =   6240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Appearance      =   3
      AutoMask        =   0   'False
      Enabled         =   0   'False
      Caption         =   "&Ok"
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
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ChkRkitchen_Click()
Priceonkitchenprint = Val(ChkRkitchen)

End Sub

Private Sub cmbDatabase_Change()
    Select Case Trim(cmbDatabase.Text)
        Case ""
            cmdForms(2).Enabled = False
        Case Else
            cmdForms(2).Enabled = True
    End Select
End Sub

Private Sub cmbDatabase_GotFocus()
    cmbDatabase.SelStart = 0
    cmbDatabase.SelLength = Len(cmbDatabase.Text)
End Sub
Private Sub cmbDatabase_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub

Private Sub cmdForms_Click(Index As Integer)
    Dim databaseE As Boolean
    databaseE = False
    On Error Resume Next
    Select Case cmdForms(Index).Caption
        Case "&Ok"
            Swiss_Round = Val(txtSwiss.Text & "")
            Slip_Printer = cmbPrinter.Text
            Server.SQL_Name = txtserver.Text
            Server.SQL_Database = cmbDatabase.Text
            Server.SQL_Password = txtPassword.Text
            Server.SQL_User = txtLogin.Text
            Workstation_No = txtWNumber.Text
            Workstation_Name = txtWork_Name.Text
            Vat_No = txtVat.Text
            ActiveUpdateServer "Delete from Xtra"
            
            ActiveUpdateServer "Insert Into Xtra (Redprintdepartment) Values ('" & CmbRed.Text & "')"
            ActiveUpdateServer "Insert Into Xtra (Redprintdepartment) Values ('" & CmbRed2.Text & "')"
            If CmbRed.Text = "" Then SaveSetting Trim(gblApp_Name), "Redprint", "Redprintenabled", "0"
            If CmbRed.Text <> "" Then SaveSetting Trim(gblApp_Name), "Redprint", "Redprintenabled", "1"
            ActiveUpdateServer "UPDATE Branch_Details set Conversion_description = '" & txtCurr.Text & "', Conversion_rate ='" & txtCurRate.Text & "' where branch_no ='" & Branch_No & "'"
            ActiveUpdateServer "Delete from Cost_Code"
            ActiveUpdateServer "Insert Into Cost_Code (One,Two,Three,Four,Five,Six,Seven,Eight,Nine,Zero) Values ('" & grdCost.TextMatrix(1, 0) & "','" & grdCost.TextMatrix(1, 1) & "','" & grdCost.TextMatrix(1, 2) & "','" & grdCost.TextMatrix(1, 3) & "','" & grdCost.TextMatrix(1, 4) & "','" & grdCost.TextMatrix(1, 5) & "','" & grdCost.TextMatrix(1, 6) & "','" & grdCost.TextMatrix(1, 7) & "','" & grdCost.TextMatrix(1, 8) & "','" & grdCost.TextMatrix(1, 9) & "')"
            If Trim(Server.SQL_Name) = "" Then
                frmSplash.Show
            Else
                Select Case opbSlip(0).Value
                    Case False
                        If opbSlip(1).Value = 0 Then
                            Slip_Printer_Type = 2
                        Else
                            Slip_Printer_Type = 1
                        End If
                    Case True: Slip_Printer_Type = 0
                End Select
                Select Case cmbPort.Text
                    Case "Auto": Slip_PrinterPort = 0
                    Case "Com1": Slip_PrinterPort = 1
                    Case "Com2": Slip_PrinterPort = 2
                    Case "Com3": Slip_PrinterPort = 3
                    Case "Com4": Slip_PrinterPort = 4
                    Case "Com5": Slip_PrinterPort = 5
                    Case "Com6": Slip_PrinterPort = 6
                    Case "Com7": Slip_PrinterPort = 7
                    Case "Com8": Slip_PrinterPort = 8
                End Select
                Kitchen_Con = Abs(chkCon.Value)
                SaveSetting Trim(gblApp_Name), "Server", "Server", txtserver.Text
                SaveSetting Trim(gblApp_Name), "Server", "SQL_User", txtLogin.Text
                SaveSetting Trim(gblApp_Name), "Server", "SQL_Password", txtPassword.Text
                SaveSetting Trim(gblApp_Name), "Server", "SQL_Database", cmbDatabase.Text
                SaveSetting Trim(gblApp_Name), "Logs", "Main_Log", txtUserLog.Text
                SaveSetting Trim(gblApp_Name), "Logs", "Error_Log", txtErrorLog.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Number", Workstation_No
                SaveSetting Trim(gblApp_Name), "Workstation", "Slip_Printer", Trim(Slip_Printer)
                SaveSetting Trim(gblApp_Name), "Workstation", "Slip_PrinterPort", Slip_PrinterPort
                SaveSetting Trim(gblApp_Name), "Workstation", "Slip_Printer_Type", Trim(Slip_Printer_Type)
                SaveSetting Trim(gblApp_Name), "Workstation", "Kitchen_Printer", cmbKitchen.ListIndex
                SaveSetting Trim(gblApp_Name), "Workstation", "Label_Printer", cmbLabel.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Label_Height", txtLabel.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Label_Width", txtwidth.Text
               
                If cmbLoc.Text = "" Then
                    MsgBox "Please set the location where the workstation is situated!", vbCritical, "HeroPOS"
                    Exit Sub
                Else
                    SaveSetting Trim(gblApp_Name), "Workstation", "Location", Val(Mid(cmbLoc.Text, 1, InStr(cmbLoc.Text, "-") - 1))
                End If
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc10", Abs(chkDiscount(0).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc20", Abs(chkDiscount(1).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc30", Abs(chkDiscount(2).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc40", Abs(chkDiscount(3).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc50", Abs(chkDiscount(4).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc60", Abs(chkDiscount(5).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc70", Abs(chkDiscount(6).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc80", Abs(chkDiscount(7).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Disc90", Abs(chkDiscount(8).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "ReplicationServ", Abs(chkReplicate.Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "AskLoc", Abs(chkAskLoc.Value)
                
                SaveSetting Trim(gblApp_Name), "Workstation", "Drawer_One_KickString", cmbDrawer1.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Drawer_Two_KickString", cmbDrawer2.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Use_Both_Drawers", Abs(chkDrawer.Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Scale_Model", cmbScale.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Scale_Port", cmbScalePort.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Scale_Settings", cmbScaleSet.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Display_Model", cmbDisplay.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Display_Port", cmbDisplayPort.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "Display_Settings", cmbDisplaySet.Text
                SaveSetting Trim(gblApp_Name), "Workstation", "PrintBarStock", Abs(chkBarStock.Value)
                SaveSetting Trim(gblApp_Name), "PriceonKitchen", "Priceonkitchenprint", Val(Priceonkitchenprint)
                
                
                
                Devices.Drawer1KickString = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Drawer_One_KickString", Default:="<Not Installed>")
                Devices.Drawer2KickString = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Drawer_Two_KickString", Default:="<Not Installed>")
                Devices.TwoDrawer = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Use_Both_Drawers", Default:=0)
                Devices.ScaleModel = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Scale_Model", Default:="<Not Installed>")
                Devices.ScalePort = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Scale_Port", Default:="<Not Set>")
                Devices.ScaleSet = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Scale_Settings", Default:="<Not Set>")
                Devices.DisplayModel = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Display_Model", Default:="<Not Installed>")
                Devices.DisplayPort = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Display_Port", Default:="<Not Set>")
                Devices.DisplaySet = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Display_Settings", Default:="<Not Set>")
                
                Workstation.Disc10 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc10", Default:=1)
                Workstation.Disc20 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc20", Default:=1)
                Workstation.Disc30 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc30", Default:=1)
                Workstation.Disc40 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc40", Default:=1)
                Workstation.Disc50 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc50", Default:=1)
                Workstation.Disc60 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc60", Default:=1)
                Workstation.Disc70 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc70", Default:=1)
                Workstation.Disc80 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc80", Default:=1)
                Workstation.Disc90 = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="Disc90", Default:=1)
                Workstation.DiscFree = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="DiscFree", Default:=1)
                ReplicationServ = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="ReplicationServ", Default:=0)
                AskLog = GetSetting(appname:=Trim(gblApp_Name), Section:="Workstation", key:="AskLoc", Default:=0)
                
                SaveSetting Trim(gblApp_Name), "Workstation", "DiscFree", Abs(chkDiscount(9).Value)
                SaveSetting Trim(gblApp_Name), "Workstation", "Display_Settings", frmSettings.cmbDisplaySet.Text
                PrintZeroItems = Abs(chkPrintZero.Value)
                PrintVoids = Abs(chkVoids.Value)
                PrintSlipTransfers = Abs(chkTransfers.Value)
                PrintBarStock = Abs(chkBarStock.Value)
                Devices.Label_Printer = cmbLabel.Text
                Devices.Barcode_Height = txtLabel.Text
                Devices.Label_Width = txtwidth.Text
                Select Case chkSOH.Value
                    Case True
                        SaveSetting Trim(gblApp_Name), "Workstation", "Stock_on_Hand", 1
                        WorkstationSOH = 1
                    Case False
                        SaveSetting Trim(gblApp_Name), "Workstation", "Stock_on_Hand", 0
                        WorkstationSOH = 0
                End Select
                Location_No = cmbLoc.ListIndex + 1
                Kitchen_Printer_No = cmbKitchen.ListIndex
                
                ActiveUpdateServer "If not exists (Select Workstation_No from Workstations WHERE Workstation_No=" & Workstation_No & ")" & _
                " Begin " & _
                "INSERT INTO [Workstations]([Workstation_No], [Workstation_Name],Location_No) values (" & Workstation_No & ",'" & Trim(Workstation_Name) & "'," & Location_No & ")" & _
                " End " & _
                " Else " & _
                " Begin " & _
                "UPDATE [Workstations] SET [Workstation_Name]='" & Trim(Workstation_Name) & "',Location_No = " & Location_No & " WHERE Workstation_No=" & Workstation_No & _
                " End"
                Select Case opbDept(0).Value
                    Case True: Dept_Order = 0
                    Case False: Dept_Order = 1
                End Select
                Select Case chkAccess.Value
                    Case False: System_Access = 0
                    Case True: System_Access = 1
                    Case Else: System_Access = 0
                End Select
                Select Case chkQCash.Value
                    Case False: QCash = 0
                    Case True: QCash = 1
                    Case Else: QCash = 0
                End Select
                Select Case chkChargePrint.Value
                    Case False: ChargeSlip = 0
                    Case True: ChargeSlip = 1
                    Case Else: ChargeSlip = 0
                End Select
                Select Case chkService.Value
                    Case False: System_Service = 0
                    Case True: System_Service = 1
                    Case Else: System_Service = 0
                End Select
                Select Case chkVoidReason.Value
                    Case False: VoidReasons = 0
                    Case True: VoidReasons = 1
                    Case Else: VoidReasons = 0
                End Select
                Select Case chkBarcode.Value
                    Case False: StockBarcode = 0
                    Case True: StockBarcode = 1
                    Case Else: StockBarcode = 0
                End Select
                Select Case chkTradePrint.Value
                    Case False: TradePrint = 0
                    Case True: TradePrint = 1
                    Case Else: TradePrint = 0
                End Select
                Select Case chkMember.Value
                    Case False: Member_No = 0
                    Case True: Member_No = 1
                    Case Else: Member_No = 0
                End Select
                Select Case chkZero.Value
                    Case False: Zero_Print = 0
                    Case True: Zero_Print = 1
                    Case Else: Zero_Print = 0
                End Select
                Select Case chkRA.Value
                    Case False: RAPrint = 0
                    Case True: RAPrint = 1
                    Case Else: RAPrint = 0
                End Select
                Select Case chkPayout.Value
                    Case False: PayoutPrint = 0
                    Case True: PayoutPrint = 1
                    Case Else: PayoutPrint = 0
                End Select
                Select Case chkCharge.Value
                    Case False: ChargePrint = 0
                    Case True: ChargePrint = 1
                    Case Else: ChargePrint = 0
                End Select
                ActiveUpdateServer "Delete from Printer_Links"
                DoEvents
                For i = 1 To grdStock.Rows - 1
                        If grdStock.TextMatrix(i, 1) = "" Or grdStock.TextMatrix(i, 1) = "<Not Linked>" Then
                            Loc_No = 0
                        Else
                            Loc_No = Val(Trim(Mid(grdStock.TextMatrix(i, 1), 1, InStr(grdStock.TextMatrix(i, 1), "-") - 1)))
                        End If
                        If grdStock.TextMatrix(i, 2) = "" Or grdStock.TextMatrix(i, 2) = "<Not Linked>" Then
                            Sales_Loc_No = 0
                        Else
                            Sales_Loc_No = Val(Trim(Mid(grdStock.TextMatrix(i, 2), 1, InStr(grdStock.TextMatrix(i, 2), "-") - 1)))
                        End If
                        ActiveUpdateServer "Insert into Printer_Links (Printer,Location_No,Sales_Location_No) values ('" & grdStock.TextMatrix(i, 0) & "'," & Loc_No & "," & Sales_Loc_No & ")"
                Next i
                ActiveUpdateServer "Update Branch_Details set Zero_Print = " & Zero_Print & ",Member_No=" & Member_No & ",Trade_Print = " & TradePrint & ",RAPrint=" & RAPrint & ",ChargePrint=" & ChargePrint & ",PayoutPrint=" & PayoutPrint & ",Stock_Barcode = " & Val(StockBarcode) & ",Void_Reasons = " & Val(VoidReasons) & ",Swiss_Round = " & Val(Swiss_Round) & ",Vat_No = '" & txtVat.Text & "',Dept_Order = " & Dept_Order & ",System_Access = " & System_Access & ",QCash = " & QCash & ",ChargeSlip = " & ChargeSlip & ",System_Service = " & System_Service & ", Kitchen_Con=" & Kitchen_Con & " where Branch_No= " & frmDetails.txtNo.Text
                If txtserver.Text <> txtserver.Tag Then
                    MsgBox "You will have to Restart the application for these settings to take Effect", vbCritical, "HeroPOS Message"
                    End
                End If
                If cmbDatabase.Text <> cmbDatabase.Tag Then
                    MsgBox "You will have to Restart the application for these settings to take Effect", vbCritical, "HeroPOS Message"
                    End
                End If
                Unload Me
                frmSplash.Show
            End If
        Case "&Cancel"
            Unload Me
            If Trim(Server.SQL_Name) = "" Then
                frmSplash.Show
            End If
            If cnnMain.State = 0 Then
                frmSplash.Show
            End If
        Case "&Help"
        Case "&Test Connection"
            Screen.MousePointer = 11
            Openconnection 10, txtserver.Text, txtLogin.Text, txtPassword.Text, cmbDatabase.Text
            Screen.MousePointer = 0
            MsgBox "Connection to Server Established...", vbOKOnly, "Server Connection"
            ActiveReadServer "Exec sp_databases"
            While Not rs.EOF
                cmbDatabase.AddItem rs.Fields("Database_Name")
                If UCase(rs.Fields("Database_Name")) = UCase(Trim(gblApp_Name)) Then
                    databaseE = True
                End If
                rs.MoveNext
            Wend
            rs.Close
'            If databaseE = False Then
'                Query = "Create Database " & Trim(gblApp_Name)
'                ActiveUpdateServer Query
'                cmbDatabase.AddItem Trim(gblApp_Name)
'                cmbDatabase.Text = Trim(gblApp_Name)
'                cnnMain.Close
'                Openconnection 30, txtServer.Text, txtLogin.Text, txtPassword.Text, cmbDatabase.Text
'            End If
    End Select
    Screen.MousePointer = 0
    On Error GoTo 0
End Sub
Private Sub cmdTest_Click(Index As Integer)
    Dim KickString As Variant
    Select Case Index
        Case 0: Drawer = 1
        Case 1: Drawer = 2
    End Select
    Dim x As Printer
    On Error GoTo trap
    If Panel_no = 2 Then frmBar.Label1 = ""
    DoEvents
    PrintErr = 0
    Slip_Port = ""
    If Trim(Slip_Printer) = "" Or Slip_Printer = "<None>" Then
        If Panel_no = 2 Then frmBar.Label1 = "No Slip Printer"
        Exit Sub
    End If
    filenum = FreeFile
    Close #filenum
    For Each x In Printers
        If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
            Slip_Port = x.Port
            Exit For
        End If
    Next
    If Slip_Port = "" Then
        On Error GoTo 0
        Exit Sub
    End If
    Open Slip_Port For Output As #filenum
    DoEvents
    If PrintErr = 1 Then
        Open Slip_Printer For Output As #filenum
        DoEvents
    End If
    Select Case Drawer
        Case 1
            KickString = Split(cmbDrawer1.Text, ",")
            If UBound(KickString) = 0 Then Print #filenum, Chr(KickString(0));
            If UBound(KickString) = 1 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1));
            If UBound(KickString) = 2 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2));
            If UBound(KickString) = 3 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2)) & Chr(KickString(3));
        Case 2
            KickString = Split(cmbDrawer1.Text, ",")
            If UBound(KickString) = 0 Then Print #filenum, Chr(KickString(0));
            If UBound(KickString) = 1 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1));
            If UBound(KickString) = 2 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2));
            If UBound(KickString) = 3 Then Print #filenum, Chr(KickString(0)) & Chr(KickString(1)) & Chr(KickString(2)) & Chr(KickString(3));
    End Select
    Print #filenum, ""
    Close #filenum
    DoEvents
    On Error GoTo 0
    Exit Sub
trap:
    If PrintErr = 1 Then
        Close #filenum
        On Error GoTo 0
        Exit Sub
    End If
    PrintErr = 1
    If Panel_no = 2 Then frmBar.Label1 = "Drawer Error"
    Close #filenum
    Resume Next
End Sub

Private Sub cmdTestLabel_Click()
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
 
 Dim byBuf(50000) As Byte
    Dim tempStr As String
    Dim Count As Integer
    Dim wwww As Integer
    tempStr = ""
    Count = 0
    wwww = (Devices.Label_Width / 5)


    tempStr = tempStr + "CB" + Chr(13)
    tempStr = tempStr + "SS3" + Chr(13)
    tempStr = tempStr + "SD20" + Chr(13)
    tempStr = tempStr + "SOT" + Chr(13)
    tempStr = tempStr + "SW400" + Chr(13)
    tempStr = tempStr + "SL" & Devices.Barcode_Height & ",20,G" + Chr(13)
    tempStr = tempStr + "T80,6,2,0,0,0,0,N,N,'Testlabel'" + Chr(13)
    tempStr = tempStr + "T80,110,3,0,0,0,0,N,N,'1234.00'" + Chr(13)
    tempStr = tempStr + "T80,70,1,0,0,0,0,N,N,'1'" + Chr(13)
    tempStr = tempStr + "B1" & wwww & ",38,1,2,3,50,0,1,'112233445566'" + Chr(13)
    tempStr = tempStr + "P1" + Chr(13)




    'Converting string to byte type
    Count = Count + UniStringToByte(tempStr, byBuf, Count)
        
    nRet = DirectWrite(DrvName, byBuf(0), Count)
    If nRet <> SEM_SUCCESS Then
        ErrorMessage (nRet)
    End If
End Sub

Private Sub Form_Activate()
    
    On Error Resume Next
    SSTab1.Tab = 0
    grdCost.TextMatrix(0, 0) = "1"
    grdCost.TextMatrix(0, 1) = "2"
    grdCost.TextMatrix(0, 2) = "3"
    grdCost.TextMatrix(0, 3) = "4"
    grdCost.TextMatrix(0, 4) = "5"
    grdCost.TextMatrix(0, 5) = "6"
    grdCost.TextMatrix(0, 6) = "7"
    grdCost.TextMatrix(0, 7) = "8"
    grdCost.TextMatrix(0, 8) = "9"
    grdCost.TextMatrix(0, 9) = "0"
    grdCost.TextMatrix(1, 0) = Cost_Code.One
    grdCost.TextMatrix(1, 1) = Cost_Code.Two
    grdCost.TextMatrix(1, 2) = Cost_Code.Three
    grdCost.TextMatrix(1, 3) = Cost_Code.Four
    grdCost.TextMatrix(1, 4) = Cost_Code.Five
    grdCost.TextMatrix(1, 5) = Cost_Code.Six
    grdCost.TextMatrix(1, 6) = Cost_Code.Seven
    grdCost.TextMatrix(1, 7) = Cost_Code.Eight
    grdCost.TextMatrix(1, 8) = Cost_Code.Nine
    grdCost.TextMatrix(1, 9) = Cost_Code.Ten
    
    chkDiscount(0).Value = Workstation.Disc10
    chkDiscount(1).Value = Workstation.Disc20
    chkDiscount(2).Value = Workstation.Disc30
    chkDiscount(3).Value = Workstation.Disc40
    chkDiscount(4).Value = Workstation.Disc50
    chkDiscount(5).Value = Workstation.Disc60
    chkDiscount(6).Value = Workstation.Disc70
    chkDiscount(7).Value = Workstation.Disc80
    chkDiscount(8).Value = Workstation.Disc90
    chkDiscount(9).Value = Workstation.DiscFree
    chkReplicate.Value = ReplicationServ
    chkAskLoc.Value = AskLog
    chkRA.Value = RAPrint
    chkPayout.Value = PayoutPrint
    chkCharge.Value = ChargePrint
    cmbDrawer1.Clear
    cmbDrawer1.AddItem "<Not Installed>"
    cmbDrawer1.AddItem "27,112,48"
    cmbDrawer1.AddItem "27,112,49"
    cmbDrawer1.Text = Devices.Drawer1KickString
    cmbDrawer2.Clear
    cmbDrawer2.AddItem "<Not Installed>"
    cmbDrawer2.AddItem "27,112,48"
    cmbDrawer2.AddItem "27,112,49"
    cmbDrawer2.Text = Devices.Drawer2KickString
    chkDrawer = Devices.TwoDrawer
    
    cmbScale.Clear
    cmbScale.AddItem "<Not Installed>"
    cmbScale.AddItem "Tereoka DS-640"
    cmbScale.AddItem "Tereoka DS-788"
    cmbScale.AddItem "Tereoka DS-860"
    cmbScale.Text = Devices.ScaleModel
    
    cmbDisplay.Clear
    cmbDisplay.AddItem "<Not Installed>"
    cmbDisplay.AddItem "Posiflex"
    cmbDisplay.AddItem "CD7220"
    cmbDisplay.Text = Devices.DisplayModel
    
    cmbScalePort.Clear
    cmbScalePort.AddItem "<Not Set>"
    cmbScalePort.AddItem "Com1"
    cmbScalePort.AddItem "Com2"
    cmbScalePort.AddItem "Com3"
    cmbScalePort.AddItem "Com4"
    cmbScalePort.AddItem "Com5"
    cmbScalePort.AddItem "Com6"
    cmbScalePort.AddItem "Com7"
    cmbScalePort.AddItem "Com8"
    cmbScalePort.AddItem "Com9"
    cmbScalePort.AddItem "Com10"
    cmbScalePort.Text = Devices.ScalePort
    
    cmbDisplayPort.Clear
    cmbDisplayPort.AddItem "<Not Set>"
    cmbDisplayPort.AddItem "Com1"
    cmbDisplayPort.AddItem "Com2"
    cmbDisplayPort.AddItem "Com3"
    cmbDisplayPort.AddItem "Com4"
    cmbDisplayPort.AddItem "Com5"
    cmbDisplayPort.AddItem "Com6"
    cmbDisplayPort.AddItem "Com7"
    cmbDisplayPort.AddItem "Com8"
    cmbDisplayPort.AddItem "Com9"
    cmbDisplayPort.AddItem "Com10"
    cmbDisplayPort.Text = Devices.DisplayPort
    
    grdStock.Rows = 1
    cmbScaleSet.Clear
    cmbScaleSet.AddItem "<Not Set>"
    cmbScaleSet.AddItem "9600,N,8,1"
    cmbScaleSet.AddItem "9600,O,8,2"
    cmbScaleSet.AddItem "14400,N,8,1"
    cmbScaleSet.AddItem "14400,0,8,2"
    cmbScaleSet.AddItem "19200,N,8,1"
    cmbScaleSet.AddItem "19200,0,8,2"
    cmbScaleSet.Text = Devices.ScaleSet
    
    cmbDisplaySet.Clear
    cmbDisplaySet.AddItem "<Not Set>"
    cmbDisplaySet.AddItem "9600,N,8,1"
    cmbDisplaySet.AddItem "9600,O,8,2"
    cmbDisplaySet.AddItem "14400,N,8,1"
    cmbDisplaySet.AddItem "14400,0,8,2"
    cmbDisplaySet.AddItem "19200,N,8,1"
    cmbDisplaySet.AddItem "19200,0,8,2"
    
    
    cmbDisplaySet.Text = Devices.DisplaySet
    
    grdStock.TextMatrix(0, 0) = "Kitchen Printer"
    grdStock.TextMatrix(0, 1) = "Stock Deduction Location"
    grdStock.TextMatrix(0, 2) = "Sales Location"
    grdStock.ColWidth(0) = 3000
    grdStock.ColWidth(1) = 3100
    
    ActiveReadServer "SELECT Kitchen1 " & _
    "FROM Products where Kitchen1<>'<None>' and Kitchen1 <>'' " & _
    "GROUP BY Kitchen1 " & _
    "Union " & _
    "SELECT Kitchen2 as Kitchen1 " & _
    "FROM Products where Kitchen2<>'<None>' and Kitchen2 <>'' " & _
    "GROUP BY Kitchen2 order by Kitchen1"
    
    While Not rs.EOF
        grdStock.Rows = grdStock.Rows + 1
        grdStock.TextMatrix(grdStock.Rows - 1, 0) = rs.Fields("Kitchen1")
        
        ActiveReadServer1 "Select Location_No,(Select Loc_Name from Locations where Locations.Location_No = Printer_Links.Location_No) as Loc_Name from Printer_Links where Printer = '" & Trim(rs.Fields("Kitchen1")) & "'"
        If rs1.RecordCount > 0 Then
            grdStock.TextMatrix(grdStock.Rows - 1, 1) = rs1.Fields("Location_No") & " - " & rs1.Fields("Loc_Name")
        Else
            grdStock.TextMatrix(grdStock.Rows - 1, 1) = "<Not Linked>"
        End If
        rs1.Close
        
        ActiveReadServer1 "Select Sales_Location_No,(Select Loc_Name from Locations where Locations.Location_No = Printer_Links.Sales_Location_No) as Sales_Loc_Name from Printer_Links where Printer = '" & Trim(rs.Fields("Kitchen1")) & "'"
        If Val(rs1.Fields("Sales_Location_No") & "") <> 0 Then
            grdStock.TextMatrix(grdStock.Rows - 1, 2) = rs1.Fields("Sales_Location_No") & " - " & rs1.Fields("Sales_Loc_Name")
        Else
            grdStock.TextMatrix(grdStock.Rows - 1, 2) = "<Not Linked>"
        End If
        rs1.Close

        rs.MoveNext
    Wend
    rs.Close
    
    cmbPort.Clear
    cmbPort.AddItem "Auto"
    cmbPort.AddItem "Com1"
    cmbPort.AddItem "Com2"
    cmbPort.AddItem "Com3"
    cmbPort.AddItem "Com4"
    cmbPort.AddItem "Com5"
    cmbPort.AddItem "Com6"
    cmbPort.AddItem "Com7"
    cmbPort.AddItem "Com8"
    Select Case Slip_PrinterPort
        Case 0: cmbPort.Text = "Auto"
        Case 1: cmbPort.Text = "Com1"
        Case 2: cmbPort.Text = "Com2"
        Case 3: cmbPort.Text = "Com3"
        Case 4: cmbPort.Text = "Com4"
        Case 5: cmbPort.Text = "Com5"
        Case 6: cmbPort.Text = "Com6"
        Case 7: cmbPort.Text = "Com7"
        Case 8: cmbPort.Text = "Com8"
    End Select
    cmbKitchen.Clear
    cmbKitchen.AddItem "Kitchen Printer #1"
    cmbKitchen.AddItem "Kitchen Printer #2"
    cmbKitchen.AddItem "Both Printers #3"
    cmbKitchen.ListIndex = Kitchen_Printer_No
    
    If Trim(Server.SQL_Name) = "" Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
    End If
    Location_Name = ""
    cmbLoc.Clear
    ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
    While Not rs.EOF
        cmbLoc.AddItem rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        If Location_No = rs.Fields("Location_No") Then
            Location_Name = rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name")
        End If
        rs.MoveNext
    Wend
    rs.Close
    cmbLoc.Text = Location_Name
    ActiveReadServer "Select isnull(Dept_Order,0) as Dept_Order,System_Access,QCash,System_Service,ChargeSlip, ISNULL(Kitchen_Con,'False') as Kitchen_Con, Conversion_description, ISNULL(Conversion_Rate, 0) as Conversion_Rate   from Branch_Details where Branch_No=" & frmDetails.txtNo
    Select Case rs.RecordCount
        Case 0
            opbDept(0).Value = 0
        Case 1
            If rs.Fields("Dept_Order") = 0 Then
                opbDept(0).Value = True
                opbDept(1).Value = False
            Else
                opbDept(0).Value = False
                opbDept(1).Value = True
            End If
    End Select
    chkCon.Value = rs.Fields("Kitchen_Con")
    Select Case Slip_Printer_Type
        Case 1
            opbSlip(0).Value = False
            opbSlip(2).Value = False
            opbSlip(1).Value = True
        Case 2
            opbSlip(2).Value = True
            opbSlip(0).Value = False
            opbSlip(0).Value = False
        Case Else
            opbSlip(0).Value = True
            opbSlip(1).Value = False
            opbSlip(2).Value = False
    End Select
    chkSOH.Value = WorkstationSOH
    Select Case rs.Fields("System_Service")
        Case 1: chkService.Value = True
        Case 0: chkService.Value = False
        Case Else: chkService.Value = False
    End Select
    Select Case rs.Fields("System_Access")
        Case 1: chkAccess.Value = True
        Case 0: chkAccess.Value = False
        Case Else: chkAccess.Value = False
    End Select
    Select Case rs.Fields("QCash")
        Case 1: chkQCash.Value = True
        Case 0: chkQCash.Value = False
        Case Else: chkQCash.Value = False
    End Select
    Select Case rs.Fields("ChargeSlip")
        Case 1: chkChargePrint.Value = True
        Case 0: chkChargePrint.Value = False
        Case Else: chkChargePrint.Value = False
    End Select
    Select Case TradePrint
        Case 1: chkTradePrint.Value = True
        Case 0: chkTradePrint.Value = False
        Case Else: chkTadePrint.Value = False
    End Select
    
    Me.txtCurr.Text = rs.Fields("Conversion_description")
    Me.txtCurRate.Text = Format(rs.Fields("Conversion_rate"), "0.000")
    rs.Close
    
    Select Case VoidReasons
        Case 0: chkVoidReason.Value = False
        Case 1: chkVoidReason.Value = True
        Case Else: chkVoidReason.Value = False
    End Select
    Select Case StockBarcode
        Case 0: chkBarcode.Value = False
        Case 1: chkBarcode.Value = True
        Case Else: chkBarcode.Value = False
    End Select
    Select Case Member_No
        Case 0: chkMember.Value = False
        Case 1: chkMember.Value = True
        Case Else: chkMember.Value = False
    End Select
    Select Case Zero_Print
        Case 0: chkZero.Value = False
        Case 1: chkZero.Value = True
        Case Else: chkZero.Value = False
    End Select
    txtWNumber.Text = Workstation_No
    txtWork_Name.Text = Trim(Workstation_Name)
    txtVat.Text = Vat_No
    cmbPrinter.AddItem "<None>"
    cmbLabel.Clear
    cmbLabel.AddItem "<None>"
    For Each x In Printers
        cmbPrinter.AddItem x.DeviceName
        cmbLabel.AddItem x.DeviceName
    Next
    cmbPrinter.AddItem "<A4 Wide Invoice>"
    cmbPrinter.AddItem "<Choose Printer>"
    cmbLabel.Text = Devices.Label_Printer
    txtLabel.Text = Devices.Barcode_Height
    txtwidth.Text = Devices.Label_Width
    cmbPrinter.Text = Slip_Printer
    If Trim(Slip_Printer) = "" Then cmbPrinter.Text = "<None>"
    chkPrintZero.Value = PrintZeroItems
    chkVoids.Value = PrintVoids
    chkTransfers.Value = PrintSlipTransfers
    chkBarStock.Value = PrintBarStock
    txtSwiss.Text = Format(Swiss_Round, "0.00")
    txtLabel.Text = Devices.Barcode_Height
    
    
    ActiveReadServer " Select Department_no from Departments"
    While Not rs.EOF
    If Len(rs.Fields("Department_no")) > 1 Then
    CmbRed.AddItem rs.Fields("Department_no")
    CmbRed2.AddItem rs.Fields("Department_no")
    End If
    rs.MoveNext
    Wend
    rs.Close
    
    ChkRkitchen.Value = Val(Priceonkitchenprint)
    ActiveReadServer " Select Redprintdepartment from Xtra"
    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
    CmbRed.Text = rs.Fields("Redprintdepartment")
    Redprintdept = rs.Fields("Redprintdepartment")
    
    If rs.RecordCount > 1 Then
    rs.MoveNext
    CmbRed2.Text = rs.Fields("Redprintdepartment")
    Redprintdept2 = rs.Fields("Redprintdepartment")
    Exit Sub
    End If
    
    
    Next i
    End If
    On Error GoTo 0
    
    
End Sub
Private Sub Form_Load()
    txtserver.Text = Trim(Server.SQL_Name)
    txtserver.Tag = Trim(Server.SQL_Name)
    txtLogin.Text = Trim(Server.SQL_User)
    txtPassword.Text = Trim(Server.SQL_Password)
    txtUserLog.Text = Trim(LogFiles.MainLog)
    txtErrorLog.Text = Trim(LogFiles.ErrorLog)
    If Trim(Server.SQL_Database) <> "" Then
        cmbDatabase.AddItem Trim(Server.SQL_Database)
        cmbDatabase.Text = Trim(Server.SQL_Database)
        cmbDatabase.Tag = Trim(Server.SQL_Database)
    End If
End Sub
Private Sub grdCost_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
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
Private Sub grdStock_EnterCell()
    If grdStock.Col = 0 Then grdStock.Col = 1
    Select Case grdStock.Col
        Case 1
            grdStock.ColComboList(1) = "<Not Linked>|"
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            While Not rs.EOF
                grdStock.ColComboList(1) = grdStock.ColComboList(1) & rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name") & "|"
                rs.MoveNext
            Wend
            rs.Close
            grdStock.Editable = flexEDKbdMouse
        Case 2
            grdStock.ColComboList(2) = "<Not Linked>|"
            ActiveReadServer "Select Location_No,Loc_Name from Locations order by Location_no"
            While Not rs.EOF
                grdStock.ColComboList(2) = grdStock.ColComboList(2) & rs.Fields("Location_No") & " - " & rs.Fields("Loc_Name") & "|"
                rs.MoveNext
            Wend
            rs.Close
            grdStock.Editable = flexEDKbdMouse
    End Select
End Sub



Private Sub txtErrorLog_GotFocus()
    txtErrorLog.SelStart = 0
    txtErrorLog.SelLength = Len(txtErrorLog.Text)
End Sub
Private Sub txtErrorLog_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtLogin_GotFocus()
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub
Private Sub txtLogin_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtServer_Change()
    If Trim(txtserver.Text) = "" Then
        cmdForms(3).Enabled = False
    Else
        cmdForms(3).Enabled = True
    End If
End Sub

Private Sub txtServer_GotFocus()
    txtserver.SelStart = 0
    txtserver.SelLength = Len(txtserver.Text)
End Sub
Private Sub txtServer_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 32
            KeyAscii = 95
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub

Private Sub txtSwiss_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If InStr(txtSwiss.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtUserLog_GotFocus()
    txtUserLog.SelStart = 0
    txtUserLog.SelLength = Len(txtUserLog.Text)
End Sub
Private Sub txtUserLog_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
    End Select
End Sub
Private Sub txtVat_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 32
            KeyAscii = 95
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
End Sub

Private Sub txtWNumber_KeyPress(KeyAscii As MSForms.ReturnInteger)
       Select Case KeyAscii
        Case 8
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtWNumber_LostFocus()
    If txtWNumber.Text = "" Then txtWNumber.Text = "0"
End Sub
Private Sub txtWork_Name_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 39
            KeyAscii = Asc("`")
        Case 32
            KeyAscii = 95
        Case 97 To 122
            KeyAscii = KeyAscii - 32
    End Select
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
