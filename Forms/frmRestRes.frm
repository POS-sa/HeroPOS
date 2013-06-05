VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRestRes 
   Appearance      =   0  'Flat
   BackColor       =   &H0077582B&
   ClientHeight    =   11400
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmRestRes.frx":0000
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPatrons 
      Appearance      =   0  'Flat
      BackColor       =   &H00E4FCFC&
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   3240
      ScaleHeight     =   3015
      ScaleWidth      =   3075
      TabIndex        =   57
      Top             =   3060
      Visible         =   0   'False
      Width           =   3105
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   14
         Left            =   90
         TabIndex        =   62
         Top             =   450
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   15
         Left            =   90
         TabIndex        =   63
         Top             =   1290
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   16
         Left            =   90
         TabIndex        =   64
         Top             =   2130
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "10"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   17
         Left            =   1080
         TabIndex        =   65
         Top             =   450
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   18
         Left            =   1080
         TabIndex        =   66
         Top             =   1290
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   32
         Left            =   1080
         TabIndex        =   67
         Top             =   2130
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "12"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   33
         Left            =   2070
         TabIndex        =   68
         Top             =   450
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   34
         Left            =   2070
         TabIndex        =   69
         Top             =   1290
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   795
         Index           =   35
         Left            =   2070
         TabIndex        =   70
         Top             =   2130
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1402
         Appearance      =   3
         BackColor       =   14737632
         Caption         =   "Other"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Left            =   240
         TabIndex        =   58
         Top             =   60
         Width           =   2565
         ForeColor       =   8421504
         BackColor       =   7821355
         VariousPropertyBits=   8388627
         Caption         =   "Number of Patrons"
         Size            =   "4524;556"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.PictureBox picPatrons1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   2370
      ScaleHeight     =   1455
      ScaleWidth      =   885
      TabIndex        =   59
      Top             =   3990
      Visible         =   0   'False
      Width           =   915
      Begin MSForms.Label lblPat 
         Height          =   405
         Left            =   30
         TabIndex        =   61
         Top             =   840
         Width           =   825
         ForeColor       =   16744576
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "4"
         Size            =   "1455;714"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label lblPatDetails 
         Height          =   465
         Left            =   0
         TabIndex        =   60
         Top             =   330
         Width           =   855
         ForeColor       =   8421504
         BackColor       =   7821355
         VariousPropertyBits=   8388627
         Caption         =   "Normally Seats"
         Size            =   "1508;820"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   60
      Top             =   4500
   End
   Begin VB.PictureBox picFlash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   780
      ScaleHeight     =   1725
      ScaleWidth      =   2535
      TabIndex        =   50
      Top             =   1080
      Visible         =   0   'False
      Width           =   2565
      Begin MSForms.Label lblflash1 
         Height          =   675
         Left            =   60
         TabIndex        =   52
         Top             =   630
         Width           =   615
         ForeColor       =   8421504
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "7"
         Size            =   "1085;1191"
         FontName        =   "Webdings"
         FontHeight      =   525
         FontCharSet     =   2
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFlash 
         Height          =   975
         Left            =   600
         TabIndex        =   51
         Top             =   420
         Width           =   1725
         ForeColor       =   8421504
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Press to Move Start to 07h30"
         Size            =   "3043;1720"
         FontName        =   "Arial"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRestRes.frx":98C5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4950
      Top             =   1050
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   4230
      Top             =   1920
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   2640
      ScaleHeight     =   4245
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   6210
      Width           =   3705
      Begin MSForms.Label lblTable 
         Height          =   225
         Left            =   220
         TabIndex        =   49
         Top             =   3720
         Width           =   1605
         BackColor       =   7821355
         VariousPropertyBits=   8388627
         Size            =   "2831;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblMany 
         Height          =   225
         Left            =   220
         TabIndex        =   48
         Top             =   3030
         Width           =   1605
         BackColor       =   7821355
         VariousPropertyBits=   8388627
         Size            =   "2831;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblAt 
         Height          =   225
         Left            =   220
         TabIndex        =   47
         Top             =   2340
         Width           =   1605
         BackColor       =   7821355
         VariousPropertyBits=   8388627
         Size            =   "2831;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblWhen 
         Height          =   225
         Left            =   220
         TabIndex        =   46
         Top             =   1650
         Width           =   1605
         BackColor       =   7821355
         VariousPropertyBits=   8388627
         Size            =   "2831;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblFor 
         Height          =   225
         Left            =   225
         TabIndex        =   45
         Top             =   960
         Width           =   3225
         BackColor       =   7821355
         VariousPropertyBits=   8388627
         Size            =   "5689;397"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Image Image1 
         Height          =   105
         Left            =   -120
         Top             =   0
         Width           =   3795
         BackColor       =   16381166
         BorderStyle     =   0
         Size            =   "6694;185"
      End
      Begin MSForms.Label Label13 
         Height          =   345
         Left            =   -30
         TabIndex        =   2
         Top             =   60
         Width           =   3735
         ForeColor       =   -2147483632
         BackColor       =   16381166
         Caption         =   "Booking Details"
         Size            =   "6588;609"
         BorderStyle     =   1
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label12 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   3645
         Width           =   1665
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Size            =   "2937;556"
         BorderStyle     =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label11 
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   3390
         Width           =   945
         BackColor       =   16777215
         Caption         =   "Table:"
         Size            =   "1667;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label10 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   2955
         Width           =   1665
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Size            =   "2937;556"
         BorderStyle     =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label9 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   2700
         Width           =   945
         BackColor       =   16777215
         Caption         =   "How many?"
         Size            =   "1667;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label8 
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   2265
         Width           =   1665
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Size            =   "2937;556"
         BorderStyle     =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   2010
         Width           =   945
         BackColor       =   16777215
         Caption         =   "At:"
         Size            =   "1667;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1575
         Width           =   1665
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Size            =   "2937;556"
         BorderStyle     =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label5 
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   1320
         Width           =   945
         BackColor       =   16777215
         Caption         =   "When?"
         Size            =   "1667;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   885
         Width           =   3345
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Size            =   "5900;556"
         BorderStyle     =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   630
         Width           =   945
         BackColor       =   16777215
         Caption         =   "For:"
         Size            =   "1667;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.PictureBox picRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9465
      Left            =   6450
      ScaleHeight     =   9435
      ScaleWidth      =   6135
      TabIndex        =   18
      Top             =   1020
      Width           =   6165
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   660
         Index           =   1
         Left            =   5430
         TabIndex        =   53
         Top             =   630
         Width           =   715
         _ExtentX        =   1270
         _ExtentY        =   1164
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   8421504
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   690
         Index           =   13
         Left            =   5430
         TabIndex        =   56
         Top             =   -30
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1217
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   8421504
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid grdMain 
         Height          =   8760
         Left            =   -30
         TabIndex        =   19
         Top             =   630
         Width           =   5505
         _cx             =   9710
         _cy             =   15452
         Appearance      =   2
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16744576
         ForeColorSel    =   -2147483634
         BackColorBkg    =   15790305
         BackColorAlternate=   16381166
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   15
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   650
         RowHeightMax    =   0
         ColWidthMin     =   100
         ColWidthMax     =   780
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
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
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   645
         Index           =   2
         Left            =   5440
         TabIndex        =   54
         Top             =   8810
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1138
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   16777215
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx cmdKey 
         Height          =   690
         Index           =   12
         Left            =   -30
         TabIndex        =   55
         Top             =   -30
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1217
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   8421504
         Caption         =   "3"
         CaptionOffsetX  =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape picGrid 
         BackColor       =   &H00F3EADC&
         FillColor       =   &H00DCC29C&
         FillStyle       =   2  'Horizontal Line
         Height          =   8865
         Left            =   5460
         Top             =   540
         Width           =   705
      End
      Begin VB.Label lblDPeriod 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   330
         TabIndex        =   42
         Top             =   120
         Width           =   5415
      End
      Begin MSForms.Image Image2 
         Height          =   675
         Left            =   -60
         Top             =   -30
         Width           =   5685
         BackColor       =   16708580
         Size            =   "10028;1191"
      End
   End
   Begin VB.PictureBox picHold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3930
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   1890
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4470
      Top             =   1020
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1035
      Index           =   7
      Left            =   420
      TabIndex        =   13
      Top             =   6240
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   1826
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Patrons"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   975
      Index           =   8
      Left            =   420
      TabIndex        =   14
      Top             =   7260
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   1720
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Tables"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   975
      Index           =   9
      Left            =   420
      TabIndex        =   15
      Top             =   8220
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   1720
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   975
      Index           =   10
      Left            =   420
      TabIndex        =   16
      Top             =   9180
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   1720
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Reports"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   975
      Index           =   0
      Left            =   420
      TabIndex        =   17
      Top             =   10140
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   1720
      Appearance      =   3
      AutoMask        =   0   'False
      BackColor       =   12632256
      Caption         =   "Exit"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   21
      Left            =   12900
      TabIndex        =   20
      Top             =   4350
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Wed"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   22
      Left            =   12900
      TabIndex        =   21
      Top             =   5280
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Thu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   24
      Left            =   12900
      TabIndex        =   22
      Top             =   7140
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Sat"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   25
      Left            =   12900
      TabIndex        =   23
      Top             =   8070
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Sun"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   930
      Index           =   6
      Left            =   12900
      TabIndex        =   24
      Top             =   9000
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1640
      Appearance      =   3
      BackColor       =   12632256
      Caption         =   "6"
      CaptionOffsetX  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   23
      Left            =   12900
      TabIndex        =   25
      Top             =   6210
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Fri"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   855
      Index           =   26
      Left            =   780
      TabIndex        =   27
      Top             =   1080
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "06h00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   855
      Index           =   27
      Left            =   1650
      TabIndex        =   28
      Top             =   1080
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "07h00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   855
      Index           =   28
      Left            =   2520
      TabIndex        =   29
      Top             =   1080
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "08h00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   855
      Index           =   29
      Left            =   780
      TabIndex        =   30
      Top             =   1980
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "09h00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   855
      Index           =   30
      Left            =   1650
      TabIndex        =   31
      Top             =   1980
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "10h00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   855
      Index           =   31
      Left            =   2520
      TabIndex        =   32
      Top             =   1980
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1508
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "11h00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   750
      Index           =   3
      Left            =   900
      TabIndex        =   33
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      Appearance      =   3
      BackColor       =   12632256
      Caption         =   "3"
      CaptionOffsetX  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   750
      Index           =   4
      Left            =   2040
      TabIndex        =   34
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      Appearance      =   3
      BackColor       =   12632256
      Caption         =   "4"
      CaptionOffsetX  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   19
      Left            =   12900
      TabIndex        =   38
      Top             =   2490
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Mon"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   20
      Left            =   12900
      TabIndex        =   39
      Top             =   3420
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   14737632
      Caption         =   "Tue"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   1980
      Index           =   11
      Left            =   5430
      TabIndex        =   40
      Top             =   1020
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   3493
      Appearance      =   3
      Enabled         =   0   'False
      BackColor       =   12632256
      Caption         =   "30min"
      CaptionOffsetX  =   2
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   900
      Index           =   5
      Left            =   12900
      TabIndex        =   26
      Top             =   1560
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      Appearance      =   3
      BackColor       =   12632256
      Caption         =   "5"
      CaptionOffsetX  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.Label lblTDate 
      Height          =   285
      Left            =   12930
      TabIndex        =   43
      Top             =   1230
      Width           =   1695
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "12-01-2005"
      Size            =   "2990;503"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblBDate 
      Height          =   285
      Left            =   12960
      TabIndex        =   44
      Top             =   10050
      Width           =   1725
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "12-01-2005"
      Size            =   "3043;503"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblTimer 
      Height          =   255
      Left            =   510
      TabIndex        =   41
      Top             =   3900
      Width           =   975
      ForeColor       =   16777215
      BackColor       =   7821355
      VariousPropertyBits=   8388627
      Caption         =   "00:00:00"
      Size            =   "1720;450"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label lblDate 
      Height          =   315
      Left            =   570
      TabIndex        =   37
      Top             =   450
      Width           =   2595
      ForeColor       =   8421504
      BackColor       =   7821355
      VariousPropertyBits=   8388627
      Caption         =   "07 Jan 2006"
      Size            =   "4577;556"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8550
      TabIndex        =   36
      Top             =   450
      Width           =   6225
   End
   Begin VB.Label lblPeriod 
      BackStyle       =   0  'Transparent
      Caption         =   "Morning"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4770
      TabIndex        =   35
      Top             =   450
      Width           =   3615
   End
   Begin MSForms.Image Image4 
      Height          =   345
      Left            =   12900
      Top             =   1170
      Width           =   1785
      BackColor       =   16777215
      Size            =   "3149;609"
   End
   Begin MSForms.Image Image5 
      Height          =   345
      Left            =   12915
      Top             =   9990
      Width           =   1785
      BackColor       =   16777215
      Size            =   "3149;609"
   End
   Begin MSForms.Image Image3 
      Height          =   9465
      Left            =   12750
      Top             =   1020
      Width           =   2085
      BackColor       =   16708580
      Size            =   "3678;16695"
   End
End
Attribute VB_Name = "frmRestRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKey_Click(Index As Integer)
      If Timer3.Enabled = True Then
            If CurrentKey <> 0 Then cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
            Timer3.Enabled = False
            DoEvents
      End If
      If cmdKey(7).Caption = "Name?" And Index <> 7 Then
            Timer4.Enabled = False
            DoEvents
            cmdKey(7).BackColor = &HC0C0C0
            cmdKey(7).Caption = "Patrons"
      End If
      CurrentKey = Index
      picHold.SetFocus
      If Index < 19 Or Index > 25 Then
            cmdKey(Index).Tag = cmdKey(Index).BackColor
            cmdKey(Index).BackColor = &HC0C0FF
            Timer3.Enabled = True
      Else
            Timer3.Enabled = False
            If cmdKey(CurrentKey).Tag <> "" Then
                  cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
            End If
            For i = 19 To 25
                  If i = Index Then
                        cmdKey(Index).BackColor = &H80C0FF
                  Else
                        cmdKey(i).BackColor = &HE0E0E0
                  End If
            Next i
      End If
      DoEvents
      Select Case Index
            Case 0
                  cmdKey(0).BackColor = &HE0E0E0
                  Timer3.Enabled = False
                  DoEvents
                  If frmRestRes.Tag = "" Then
                      frmSplash.Show
                  End If
                  DoEvents
                  Me.Hide
            Case 1
                  grdMain.TopRow = grdMain.TopRow - 1
            Case 2
                   grdMain.TopRow = grdMain.TopRow + 1
            Case 3, 4
                  SwitchClock Index
            Case 5, 6
                  SwitchDate Index
            Case 12
                  grdMain.LeftCol = grdMain.LeftCol - 1
            Case 13
                  grdMain.LeftCol = grdMain.LeftCol + 1
            Case 19 To 25
                  lblDPeriod.Caption = Format(Dates(Index - 19), "DD MMM YYYY DDDD")
                  DateSelect = Dates(Index - 19)
      End Select
      Select Case cmdKey(Index).Caption
            Case "Settings"
                  frmSettings.Show vbModal
            Case "30min"
                  If grdMain.Text = "S" Then
                        If Right(StartTime, 2) = "00" Then
                              Set grdMain.CellPicture = LoadPicture(App.Path & "\Icons\\SRHH.ico")
                              StartTime = Mid(StartTime, 1, InStr(StartTime, "h")) & "30"
                              lblFlash.Caption = "Press to Move Start to " & Left(StartTime, 3) & "00"
                        Else
                              Set grdMain.CellPicture = LoadPicture(App.Path & "\Icons\\SRH.ico")
                              StartTime = Mid(StartTime, 1, InStr(StartTime, "h")) & "00"
                              lblFlash.Caption = "Press to Move Start to " & Left(StartTime, 3) & "30"
                        End If
                        lblAt.Caption = StartTime & " to " & StopTime
                  End If
                  If grdMain.Text = "ST" Then
                        If Right(StopTime, 2) = "00" Then
                              Set grdMain.CellPicture = LoadPicture(App.Path & "\Icons\\STHH.ico")
                              StopTime = Mid(grdMain.TextMatrix(0, GridXS), 1, InStr(grdMain.TextMatrix(0, GridXS), "h")) & "30"
                              If grdMain.Col = 24 Then
                                    lblFlash.Caption = "Press to Move Stop to " & "06h00"
                              Else
                                    lblFlash.Caption = "Press to Move Stop to " & Left(grdMain.TextMatrix(0, GridXS + 1), 3) & "00"
                              End If
                        Else
                              Set grdMain.CellPicture = LoadPicture(App.Path & "\Icons\\STH.ico")
                              If grdMain.Col = 24 Then
                                    StopTime = "06h00"
                              Else
                                    StopTime = grdMain.TextMatrix(0, GridXS + 1)
                              End If
                              lblFlash.Caption = "Press to Move Stop to " & Left(grdMain.TextMatrix(0, GridXS), 3) & "30"
                        End If
                        lblAt.Caption = StartTime & " to " & StopTime
                  End If
                  grdMain.SetFocus
            Case "Cancel"
                  lblPatDetails.Caption = "Normally Seats"
                  Timer4.Enabled = False
                  Timer3.Enabled = False
                  DoEvents
                  cmdKey(7).BackColor = &HE0E0E0
                  cmdKey(8).BackColor = &HE0E0E0
                  cmdKey(9).BackColor = &HE0E0E0
                  cmdKey(10).BackColor = &HE0E0E0
                  cmdKey(0).BackColor = &HE0E0E0
                  grdMain.TextMatrix(GridY, GridX) = ""
                  grdMain.TextMatrix(GridY, GridXS) = ""
                  For i = 19 To 25
                        cmdKey(i).Enabled = True
                  Next i
                  For i = 26 To 31
                        cmdKey(i).Enabled = True
                  Next i
                  cmdKey(1).Enabled = True
                  cmdKey(2).Enabled = True
                  cmdKey(5).Enabled = True
                  cmdKey(6).Enabled = True
                  cmdKey(9).Enabled = True
                  cmdKey(10).Enabled = True
                  cmdKey(11).Enabled = False
                  cmdKey(0).Enabled = True
                  picFlash.Visible = False
                  cmdKey(7).Caption = "Patrons"
                  cmdKey(8).Caption = "Tables"
                  cmdKey(9).Caption = "Settings"
                  grdMain.Col = GridX
                  grdMain.Row = GridY
                  Set grdMain.CellPicture = LoadPicture("")
                  grdMain.Col = GridXS
                  Set grdMain.CellPicture = LoadPicture("")
                  For i = GridX + 1 To GridXS - 1
                        grdMain.Col = i
                        Set grdMain.CellPicture = LoadPicture("")
                        grdMain.TextMatrix(GridY, i) = ""
                  Next i
                  grdMain.Cell(flexcpBackColor, GridY, 1, GridY, 24) = 0
                  GridX = 0
                  GridY = 0
                  lblAt.Caption = ""
                  lblFor.Caption = ""
                  lblTable.Caption = ""
                  lblWhen.Caption = ""
                  lblMany.Caption = ""
                  lblPat.Caption = ""
                  picPatrons1.Visible = False
                  picPatrons.Visible = False
                  CurrentKey = 0
            Case "Name?"
                  Timer3.Enabled = False
                  cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
                  frmKeyBoard.Show vbModal
                  If lblFor.Caption <> "" Then
                        Set grdMain.CellPicture = LoadPicture(App.Path & "\Icons\\SRH.ico")
                        GridY = grdMain.Row
                        GridX = grdMain.Col
                        StartTime = grdMain.TextMatrix(0, GridX)
                        grdMain.TextMatrix(GridY, GridX) = "S"
                        For i = 19 To 25
                              cmdKey(i).Enabled = False
                        Next i
                        For i = 26 To 31
                              cmdKey(i).Enabled = False
                        Next i
                        cmdKey(1).Enabled = False
                        cmdKey(2).Enabled = False
                        cmdKey(5).Enabled = False
                        cmdKey(6).Enabled = False
                        cmdKey(9).Enabled = False
                        cmdKey(10).Enabled = False
                        cmdKey(0).Enabled = False
                        cmdKey(11).Enabled = True
                        picFlash.Visible = True
                        cmdKey(7).Caption = "Cancel"
                        cmdKey(7).BackColor = &HE0E0E0
                        cmdKey(8).Caption = "Stop when?"
                        grdMain.Cell(flexcpBackColor, GridY, 1, GridY, 24) = &HC0FFFF
                        Timer4.Enabled = True
                        lblFlash.Caption = "Press to Move Start to " & Left(grdMain.TextMatrix(0, GridX), 3) & "30"
                        lblAt.Caption = grdMain.TextMatrix(0, GridX)
                        lblTable.Caption = Mid(grdMain.TextMatrix(GridY, 0), InStr(grdMain.TextMatrix(GridY, 0), " ") + 1)
                        lblWhen.Caption = Format(DateSelect, "DD MMM YYYY DDD")
                        lblMany.Caption = "4"
                        lblPat.Caption = "4"
                        grdMain.SetFocus
                  End If
            Case "Stop when?"
                  Timer3.Enabled = False
                  cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
                  Timer4.Enabled = True
                  cmdKey(8).Caption = "How Many?"
                  cmdKey(9).Caption = "Save"
                  cmdKey(9).Enabled = True
                  GridXS = grdMain.Col
                  If GridX = GridXS Then
                        cmdKey(11).Enabled = False
                        Set grdMain.CellPicture = LoadPicture(App.Path & "\Icons\\SH.ico")
                        lblFlash.Caption = "No Half Hour Selection Possible"
                  Else
                        Set grdMain.CellPicture = LoadPicture(App.Path & "\Icons\\STH.ico")
                        For i = GridX + 1 To GridXS - 1
                              grdMain.Col = i
                              Set grdMain.CellPicture = LoadPicture(App.Path & "\Icons\\Hour.ico")
                              grdMain.TextMatrix(GridY, i) = "H"
                        Next i
                        grdMain.Col = GridXS
                  End If
                  grdMain.TextMatrix(GridY, GridXS) = "ST"
                  lblFlash.Caption = "Press to Move Stop to " & Left(grdMain.TextMatrix(0, GridXS), 3) & "30"
                  grdMain.SetFocus
            Case "How Many?"
                  Timer3.Enabled = False
                  cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
                  Timer4.Enabled = False
                  picPatrons1.Visible = True
                  picPatrons.Visible = True
                  cmdKey(8).BackColor = &HE0E0E0
                  cmdKey(9).BackColor = &HE0E0E0
                  DoEvents
                  Timer4.Enabled = True
            Case "Save"
                  CurrentKey = 0
                  Timer3.Enabled = False
                  Timer4.Enabled = False
                  DoEvents
                  cmdKey(8).BackColor = &HE0E0E0
                  cmdKey(9).BackColor = &HE0E0E0
                  grdMain.Cell(flexcpBackColor, GridY, 1, GridY, 24) = 0
                  SaveRes
                  GridY = 0
                  GridX = 0
                  GridXS = 0
                  StartTime = ""
                  StopTime = ""
                  For i = 19 To 25
                        cmdKey(i).Enabled = True
                  Next i
                  For i = 26 To 31
                        cmdKey(i).Enabled = True
                  Next i
                  cmdKey(1).Enabled = True
                  cmdKey(2).Enabled = True
                  cmdKey(5).Enabled = True
                  cmdKey(6).Enabled = True
                  cmdKey(9).Enabled = True
                  cmdKey(10).Enabled = True
                  cmdKey(11).Enabled = False
                  cmdKey(0).Enabled = True
                  picFlash.Visible = False
                  cmdKey(7).Caption = "Patrons"
                  cmdKey(8).Caption = "Tables"
                  cmdKey(9).Caption = "Settings"
                  lblAt.Caption = ""
                  lblFor.Caption = ""
                  lblTable.Caption = ""
                  lblWhen.Caption = ""
                  lblMany.Caption = ""
                  lblPat.Caption = ""
                  picPatrons.Visible = False
                  picPatrons1.Visible = False
                  lblPatDetails.Caption = "Normally Seats"
            Case "2", "3", "4", "5", "6", "8", "10", "12"
                  If picPatrons.Visible = True Then
                        Timer4.Enabled = False
                        DoEvents
                        cmdKey(9).BackColor = &HE0E0E0
                        DoEvents
                        lblMany.Caption = cmdKey(Index).Caption
                        lblPat.Caption = cmdKey(Index).Caption
                        picPatrons.Visible = False
                        picPatrons1.Visible = False
                        lblPatDetails.Caption = "Now Seats"
                        Timer4.Enabled = True
                  End If
            Case "Other"
                  Timer3.Enabled = False
                  cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
                  frmKeyBoard.Show vbModal
                  Timer4.Enabled = False
                  DoEvents
                  cmdKey(9).BackColor = &HE0E0E0
                  DoEvents
                  If Val(lblMany.Caption) <> 0 Then
                        lblPat.Caption = lblMany.Caption
                        picPatrons.Visible = False
                        picPatrons1.Visible = False
                        lblPatDetails.Caption = "Now Seats"
                        Timer4.Enabled = True
                  End If
            Case Else
                  
      End Select
End Sub
Private Sub Form_Load()
      Load frmKeyBoard
      DateSelect = Date
      lblDPeriod.Caption = Format(Date, "DD MMM YYYY DDDD")
      grdMain.ColAlignment(0) = flexAlignCenterCenter
      lblDate.Caption = Format(Date, "DD MMM YYYY")
      For i = 1 To 18
            grdMain.TextMatrix(0, i) = Format(Trim(Str(i + 5)), "00") & "h00"
            grdMain.ColAlignment(i) = flexAlignCenterCenter
      Next i
      For i = 19 To 24
            grdMain.TextMatrix(0, i) = Format(Trim(Str(i - 19)), "00") & "h00"
            grdMain.ColAlignment(i) = flexAlignCenterCenter
      Next i
      grdMain.LeftCol = 1
      For i = 1 To 14
            grdMain.TextMatrix(i, 0) = "Tabl " & i
      Next i
      backcount = -1
      For i = 19 To 25
            If cmdKey(i).Caption = Format(Date, "DDD") Or cmdKey(i).Caption = "Today" Then
                  cmdKey(i).BackColor = &H80C0FF
                  cmdKey(i).Caption = "Today"
                  Dates(i - 19) = Date
                  backcount = i
            Else
                  If backcount <> -1 Then Dates(i - 19) = DateAdd("d", 1, Dates(i - 20))
            End If
      Next i
      If backcount <> -1 Then
            For i = backcount - 19 To 0 Step -1
                  If i <> 6 Then
                        Dates(i) = DateAdd("d", -1, Dates(i + 1))
                  End If
            Next i
      End If
      lblTDate = Format(Dates(0), "dd mmm yyyy")
      lblBDate = Format(Dates(6), "dd mmm yyyy")
End Sub
Private Sub grdMain_Click()
      DoEvents
      If cmdKey(7).Caption = "Patrons" Then
            If grdMain.TextMatrix(grdMain.Row, grdMain.Col) = "" Then
                  cmdKey(7).Caption = "Name?"
                  Timer3.Enabled = False
                  If CurrentKey <> 0 Then
                        If cmdKey(CurrentKey).Tag <> "" Then
                              cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
                        End If
                  End If
                  lblAt.Caption = ""
                  lblFor.Caption = ""
                  lblWhen.Caption = ""
                  lblTable.Caption = ""
                  lblMany.Caption = ""
                  Timer4.Enabled = True
                  Exit Sub
            End If
      End If
      If cmdKey(7).Caption = "Name?" Then
            If grdMain.TextMatrix(grdMain.Row, grdMain.Col) <> "" Then
                  If Trim(StartTime) = "" Then
                        cmdKey(7).Caption = "Patrons"
                        Timer3.Enabled = False
                        Timer4.Enabled = False
                        cmdKey(7).BackColor = &HE0E0E0
                  End If
            End If
      End If
      If cmdKey(8).Caption = "Stop when?" Then
            If grdMain.Col = 24 Then
                  lblAt.Caption = StartTime & " to 06h00"
                  StopTime = "06h00"
            Else
                  lblAt.Caption = StartTime & " to " & grdMain.TextMatrix(0, grdMain.Col + 1)
                  StopTime = grdMain.TextMatrix(0, grdMain.Col + 1)
            End If
            For i = GridX + 1 To grdMain.Col
                  If grdMain.TextMatrix(GridY, i) <> "" Then
                        If Mid(grdMain.TextMatrix(GridY, i), InStr(grdMain.TextMatrix(GridY, i), "|") + 4, 2) = "30" Then
                              grdMain.Col = i
                        Else
                              grdMain.Col = i - 1
                        End If
                        Exit For
                  End If
            Next i
      End If
      If cmdKey(8).Caption = "How Many?" Then
            cmdKey(11).Enabled = True
            If grdMain.Col = GridX Then
                  If Right(StartTime, 2) = "30" Then
                        lblFlash.Caption = "Press to Move Start to " & Left(StartTime, 3) & "00"
                  Else
                        lblFlash.Caption = "Press to Move Start to " & Left(StartTime, 3) & "30"
                  End If
             End If
             If grdMain.Col = GridXS Then
                  If Right(StopTime, 2) = "30" Then
                        If grdMain.Col = 24 Then
                              lblFlash.Caption = "Press to Move Stop to " & "06h00"
                        Else
                              lblFlash.Caption = "Press to Move Stop to " & Left(grdMain.TextMatrix(0, GridXS + 1), 3) & "00"
                        End If
                  Else
                        lblFlash.Caption = "Press to Move Stop to " & Left(grdMain.TextMatrix(0, GridXS), 3) & "30"
                   End If
             End If
             If grdMain.Col <> GridX And grdMain.Col <> GridXS Then
                  cmdKey(11).Enabled = False
                  lblFlash.Caption = "Select Start or Stop Time to Select"
             End If
             If GridX = GridXS Then
                  cmdKey(11).Enabled = False
                  lblFlash.Caption = "No Half Hour Selection Possible"
             End If
      End If
      If Left(grdMain.TextMatrix(grdMain.Row, grdMain.Col), 3) = "REC" Then
            If cmdKey(7).Caption = "Patrons" Then LoadRes
      End If
      lblInfo.Caption = "No Reservations for " & grdMain.TextMatrix(0, grdMain.Col)
End Sub
Private Sub grdMain_RowColChange()
      If cmdKey(8).Caption = "Stop when?" Then
            If grdMain.Col < GridX Then
                  grdMain.Col = GridX
            End If
            If grdMain.Row <> GridY Then
                  grdMain.Row = GridY
            End If
            If grdMain.Col > grdMain.LeftCol + 6 Then
                  grdMain.LeftCol = grdMain.Col
            End If
      End If
      If cmdKey(8).Caption = "How Many?" Then
            If grdMain.Col < GridX Then
                  grdMain.Col = GridX
            End If
            If grdMain.Col > GridXS Then
                  grdMain.Col = GridXS
            End If
            If grdMain.Row <> GridY Then
                  grdMain.Row = GridY
            End If
            If grdMain.Col > grdMain.LeftCol + 6 Then
                  grdMain.LeftCol = grdMain.Col
            End If
      End If
End Sub
Private Sub Timer1_Timer()
      lblTimer.Caption = Format(Time, "HH:MM:SS")
End Sub
Private Sub Timer2_Timer()
      lblDate.Caption = Format(Date, "DD MMM YYYY")
End Sub
Private Sub Timer3_Timer()
      Static ClockCycle As Integer
      ClockCycle = ClockCycle + 1
      Select Case cmdKey(CurrentKey).BackColor
            Case &HC0C0FF
                  If CurrentKey <> 0 Then cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
            Case Else
                  cmdKey(CurrentKey).BackColor = &HC0C0FF
      End Select
      If ClockCycle = 6 Then
            ClockCycle = 0
            If CurrentKey <> 0 Then cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
            Timer3.Enabled = False
      End If
End Sub
Private Sub SwitchClock(Action)
      DoEvents
      Select Case Action
            Case 3
                  Select Case lblPeriod.Caption
                        Case "Morning"
                              lblPeriod.Caption = "After Midnight"
                              grdMain.LeftCol = 19
                        Case "After Midnight"
                              lblPeriod.Caption = "Evening"
                              grdMain.LeftCol = 13
                        Case "Afternoon"
                              lblPeriod.Caption = "Morning"
                              grdMain.LeftCol = 1
                        Case "Evening"
                              lblPeriod.Caption = "Afternoon"
                              grdMain.LeftCol = 7
                  End Select
            Case 4
                  Select Case lblPeriod.Caption
                        Case "Morning"
                              lblPeriod.Caption = "Afternoon"
                              grdMain.LeftCol = 7
                        Case "After Midnight"
                              lblPeriod.Caption = "Morning"
                              grdMain.LeftCol = 1
                        Case "Afternoon"
                              lblPeriod.Caption = "Evening"
                              grdMain.LeftCol = 13
                        Case "Evening"
                              lblPeriod.Caption = "After Midnight"
                              grdMain.LeftCol = 19
                  End Select
      End Select
      b = grdMain.LeftCol
      For i = 26 To 31
            b = b + 1
            cmdKey(i).Caption = grdMain.TextMatrix(0, b - 1)
      Next i
      DoEvents
End Sub
Private Sub SwitchDate(Action)
      cmdKey(CurrentKey).BackColor = cmdKey(CurrentKey).Tag
      Timer3.Enabled = False
      DoEvents
      Select Case Action
            Case 5
                  For i = 25 To 20 Step -1
                        cmdKey(i).Caption = cmdKey(i - 1).Caption
                        cmdKey(i).BackColor = cmdKey(i - 1).BackColor
                        Dates(i - 19) = Dates(i - 20)
                  Next i
                  cmdKey(19).BackColor = &HE0E0E0
                  Dates(0) = DateAdd("d", -1, Dates(1))
                  cmdKey(19).Caption = Format(Dates(0), "DDD")
            Case 6
                  For i = 19 To 24
                        cmdKey(i).Caption = cmdKey(i + 1).Caption
                        cmdKey(i).BackColor = cmdKey(i + 1).BackColor
                        Dates(i - 19) = Dates(i - 18)
                  Next i
                  cmdKey(24).BackColor = &HE0E0E0
                  Dates(6) = DateAdd("d", 1, Dates(5))
                  cmdKey(25).Caption = Format(Dates(6), "DDD")
      End Select
      For i = 0 To 6
            If Dates(i) = Date Then
                 cmdKey(19 + i).Caption = "Today"
            End If
      Next i
      lblTDate = Format(Dates(0), "dd mmm yyyy")
      lblBDate = Format(Dates(6), "dd mmm yyyy")
End Sub
Private Sub Timer4_Timer()
      Static CycleCount As Integer
      CycleCount = CycleCount + 1
      If cmdKey(8).Caption = "Stop when?" Or cmdKey(8).Caption = "How Many?" Then
            If picPatrons.Visible = True Then
                  Select Case cmdKey(9).BackColor
                        Case &HFF8080
                              cmdKey(9).BackColor = &HE0E0E0
                        Case Else
                              cmdKey(9).BackColor = &HFF8080
                  End Select
            Else
                  If lblPatDetails.Caption = "Now Seats" Then
                        Select Case cmdKey(9).BackColor
                              Case &HFF8080
                                    cmdKey(9).BackColor = &HE0E0E0
                              Case Else
                                    cmdKey(9).BackColor = &HFF8080
                        End Select
                  Else
                        Select Case cmdKey(8).BackColor
                              Case &HFF8080
                                    cmdKey(8).BackColor = &HE0E0E0
                              Case Else
                                    cmdKey(8).BackColor = &HFF8080
                        End Select
                  End If
            End If
      End If
      If cmdKey(8).Caption = "Tables" Then
            Select Case cmdKey(7).BackColor
                  Case &HFF8080
                        cmdKey(7).BackColor = &HE0E0E0
                  Case Else
                        cmdKey(7).BackColor = &HFF8080
            End Select
      End If
      If CycleCount = 8 Then
            CycleCount = 0
            Timer4.Enabled = False
            cmdKey(7).BackColor = &HE0E0E0
            cmdKey(8).BackColor = &HE0E0E0
            cmdKey(9).BackColor = &HE0E0E0
      End If
End Sub
Private Sub SaveRes()
      RecordString = "REC|" & StartTime & "|" & StopTime & "|" & DateSelect & "|" & lblTable.Caption & "|" & lblMany.Caption & "|" & lblFor.Caption
      For i = GridX To GridXS
            grdMain.TextMatrix(GridY, i) = RecordString
      Next i
End Sub
Private Sub LoadRes()
      Dim ResSplit As Variant
      ResSplit = Split(grdMain.TextMatrix(grdMain.Row, grdMain.Col), "|", -1)
      lblAt.Caption = ResSplit(1) & " to " & ResSplit(2)
      lblFor.Caption = ResSplit(6)
      lblWhen.Caption = Format(ResSplit(3), "DD MMM YYYY DDD")
      lblTable.Caption = ResSplit(4)
      lblMany.Caption = ResSplit(5)
End Sub

