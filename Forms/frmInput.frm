VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmInput 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H000000C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInput.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin BTNENHLib4.BtnEnh lblTransfer 
      Height          =   1260
      Left            =   390
      TabIndex        =   64
      Top             =   2340
      Visible         =   0   'False
      Width           =   9360
      _Version        =   524298
      _ExtentX        =   16510
      _ExtentY        =   2222
      _StockProps     =   66
      Caption         =   "Show All Tabs"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   18
      Picture         =   "c:\Temp\Code\950\GUI\GUI NEW\Gold.JPG"
      BackColorContainer=   14737632
      SpecialEffect   =   2
      LogPixels       =   96
      Clickable       =   0   'False
      SpecialEffectFactor=   1
      TextureBevelFactor=   4
      UserData        =   0.1
      dibPicture      =   "frmInput.frx":E4CB
      textCaption     =   "frmInput.frx":F3F1
      textLT          =   "frmInput.frx":F46B
      textCT          =   "frmInput.frx":F483
      textRT          =   "frmInput.frx":F49B
      textLM          =   "frmInput.frx":F4B3
      textRM          =   "frmInput.frx":F4CB
      textLB          =   "frmInput.frx":F4E3
      textCB          =   "frmInput.frx":F4FB
      textRB          =   "frmInput.frx":F513
      colorBack       =   "frmInput.frx":F52B
      colorIntern     =   "frmInput.frx":F555
      colorMO         =   "frmInput.frx":F57F
      colorFocus      =   "frmInput.frx":F5A9
      colorDisabled   =   "frmInput.frx":F5D3
      colorPressed    =   "frmInput.frx":F5FD
      Style           =   7
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
      SurfaceTransparentZone=   1
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   885
      Left            =   4650
      TabIndex        =   48
      Top             =   345
      Visible         =   0   'False
      Width           =   10455
      _Version        =   524298
      _ExtentX        =   18441
      _ExtentY        =   1561
      _StockProps     =   66
      Caption         =   "Invalid Key Pressed"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   15
      BackColorContainer=   12632256
      SpecialEffect   =   3
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "frmInput.frx":F627
      textLT          =   "frmInput.frx":F6AD
      textCT          =   "frmInput.frx":F6C5
      textRT          =   "frmInput.frx":F6DD
      textLM          =   "frmInput.frx":F6F5
      textRM          =   "frmInput.frx":F70D
      textLB          =   "frmInput.frx":F725
      textCB          =   "frmInput.frx":F73D
      textRB          =   "frmInput.frx":F755
      colorBack       =   "frmInput.frx":F76D
      colorIntern     =   "frmInput.frx":F797
      colorMO         =   "frmInput.frx":F7C1
      colorFocus      =   "frmInput.frx":F7EB
      colorDisabled   =   "frmInput.frx":F815
      colorPressed    =   "frmInput.frx":F83F
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1605
      Index           =   13
      Left            =   13290
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2940
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2734
      _ExtentY        =   2831
      _StockProps     =   66
      Caption         =   "Ok"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      CornerFactor    =   30
      BackColorContainer=   11983853
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":F869
      textLT          =   "frmInput.frx":F8CD
      textCT          =   "frmInput.frx":F8E5
      textRT          =   "frmInput.frx":F8FD
      textLM          =   "frmInput.frx":F915
      textRM          =   "frmInput.frx":F92D
      textLB          =   "frmInput.frx":F945
      textCB          =   "frmInput.frx":F95D
      textRB          =   "frmInput.frx":F975
      colorBack       =   "frmInput.frx":F98D
      colorIntern     =   "frmInput.frx":F9B7
      colorMO         =   "frmInput.frx":F9E1
      colorFocus      =   "frmInput.frx":FA0B
      colorDisabled   =   "frmInput.frx":FA35
      colorPressed    =   "frmInput.frx":FA5F
      HollowFrame     =   -1  'True
      LightDirection  =   7
      ShapeHeadFactor =   40
      ShapeLineFactor =   40
   End
   Begin VB.PictureBox picSlip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9315
      Left            =   4170
      ScaleHeight     =   9285
      ScaleWidth      =   255
      TabIndex        =   50
      Top             =   1650
      Visible         =   0   'False
      Width           =   285
      Begin btButtonEx.ButtonEx cmdClose 
         Height          =   630
         Left            =   4230
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   8520
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1111
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   8421504
         Caption         =   "Close"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   630
         Index           =   1
         Left            =   5610
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   7770
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1111
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   8421504
         Caption         =   "6"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin btButtonEx.ButtonEx cmdArrow 
         Height          =   600
         Index           =   0
         Left            =   5610
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   690
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1058
         Appearance      =   3
         BackColor       =   12632256
         BorderColor     =   8421504
         Caption         =   "5"
         CaptionOffsetX  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   0
      End
      Begin VSFlex8Ctl.VSFlexGrid grdMain 
         Height          =   7680
         Left            =   90
         TabIndex        =   51
         Top             =   690
         Width           =   5535
         _cx             =   9763
         _cy             =   13547
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12.75
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
         BackColorSel    =   15523287
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16381166
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSForms.Label lblDateOpened 
         Height          =   315
         Left            =   120
         TabIndex        =   60
         Top             =   8850
         Width           =   3915
         ForeColor       =   8388608
         BackColor       =   -2147483643
         VariousPropertyBits=   8388627
         Size            =   "6906;556"
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblWorkstation 
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   8490
         Width           =   3915
         ForeColor       =   8388608
         BackColor       =   -2147483643
         VariousPropertyBits=   8388627
         Size            =   "6906;556"
         FontName        =   "Arial Narrow"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblTender 
         Height          =   375
         Left            =   4470
         TabIndex        =   56
         Top             =   130
         Width           =   1725
         ForeColor       =   7555868
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "0.00"
         Size            =   "3043;661"
         FontName        =   "Arial Narrow"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblCash 
         Height          =   435
         Left            =   240
         TabIndex        =   55
         Top             =   135
         Width           =   1965
         ForeColor       =   7555868
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Subtotal"
         Size            =   "3466;767"
         FontName        =   "Arial Narrow"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape picGrid 
         BackColor       =   &H00F3EADC&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00DCC29C&
         FillStyle       =   2  'Horizontal Line
         Height          =   7125
         Left            =   5640
         Top             =   700
         Width           =   705
      End
      Begin MSForms.Image Image6 
         Height          =   7695
         Left            =   90
         Top             =   690
         Width           =   6285
         BorderStyle     =   0
         SpecialEffect   =   2
         Size            =   "11086;13573"
      End
      Begin MSForms.Image Image7 
         Height          =   525
         Left            =   90
         Top             =   110
         Width           =   6265
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "11051;926"
      End
      Begin MSForms.Image Image5 
         Height          =   9285
         Left            =   0
         Top             =   0
         Width           =   6465
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "11404;16378"
      End
   End
   Begin BTNENHLib4.BtnEnh cmdLogoff 
      Height          =   1755
      Left            =   375
      TabIndex        =   20
      Top             =   600
      Width           =   2565
      _Version        =   524298
      _ExtentX        =   4524
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Log Off"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":FA89
      textLT          =   "frmInput.frx":FAF7
      textCT          =   "frmInput.frx":FB0F
      textRT          =   "frmInput.frx":FB27
      textLM          =   "frmInput.frx":FB3F
      textRM          =   "frmInput.frx":FB57
      textLB          =   "frmInput.frx":FB6F
      textCB          =   "frmInput.frx":FB87
      textRB          =   "frmInput.frx":FB9F
      colorBack       =   "frmInput.frx":FBB7
      colorIntern     =   "frmInput.frx":FBE1
      colorMO         =   "frmInput.frx":FC0B
      colorFocus      =   "frmInput.frx":FC35
      colorDisabled   =   "frmInput.frx":FC5F
      colorPressed    =   "frmInput.frx":FC89
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   3
      Left            =   2940
      TabIndex        =   21
      Top             =   600
      Width           =   1635
      _Version        =   524298
      _ExtentX        =   2884
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Transfer Table"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   4
      CornerFactor    =   100
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":FCB3
      textLT          =   "frmInput.frx":FD2F
      textCT          =   "frmInput.frx":FD47
      textRT          =   "frmInput.frx":FD5F
      textLM          =   "frmInput.frx":FD77
      textRM          =   "frmInput.frx":FD8F
      textLB          =   "frmInput.frx":FDA7
      textCB          =   "frmInput.frx":FDBF
      textRB          =   "frmInput.frx":FDD7
      colorBack       =   "frmInput.frx":FDEF
      colorIntern     =   "frmInput.frx":FE19
      colorMO         =   "frmInput.frx":FE43
      colorFocus      =   "frmInput.frx":FE6D
      colorDisabled   =   "frmInput.frx":FE97
      colorPressed    =   "frmInput.frx":FEC1
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   945
      Index           =   6
      Left            =   4560
      TabIndex        =   24
      Top             =   1410
      Width           =   1815
      _Version        =   524298
      _ExtentX        =   3201
      _ExtentY        =   1667
      _StockProps     =   66
      Caption         =   "Print Bill"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   20
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":FEEB
      textLT          =   "frmInput.frx":FF5F
      textCT          =   "frmInput.frx":FF77
      textRT          =   "frmInput.frx":FF8F
      textLM          =   "frmInput.frx":FFA7
      textRM          =   "frmInput.frx":FFBF
      textLB          =   "frmInput.frx":FFD7
      textCB          =   "frmInput.frx":FFEF
      textRB          =   "frmInput.frx":10007
      colorBack       =   "frmInput.frx":1001F
      colorIntern     =   "frmInput.frx":10049
      colorMO         =   "frmInput.frx":10073
      colorFocus      =   "frmInput.frx":1009D
      colorDisabled   =   "frmInput.frx":100C7
      colorPressed    =   "frmInput.frx":100F1
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   945
      Index           =   8
      Left            =   6360
      TabIndex        =   26
      Top             =   1410
      Width           =   1725
      _Version        =   524298
      _ExtentX        =   3043
      _ExtentY        =   1667
      _StockProps     =   66
      Caption         =   "Change Covers"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   20
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":1011B
      textLT          =   "frmInput.frx":10195
      textCT          =   "frmInput.frx":101AD
      textRT          =   "frmInput.frx":101C5
      textLM          =   "frmInput.frx":101DD
      textRM          =   "frmInput.frx":101F5
      textLB          =   "frmInput.frx":1020D
      textCB          =   "frmInput.frx":10225
      textRB          =   "frmInput.frx":1023D
      colorBack       =   "frmInput.frx":10255
      colorIntern     =   "frmInput.frx":1027F
      colorMO         =   "frmInput.frx":102A9
      colorFocus      =   "frmInput.frx":102D3
      colorDisabled   =   "frmInput.frx":102FD
      colorPressed    =   "frmInput.frx":10327
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdTable 
      Height          =   8370
      Left            =   0
      TabIndex        =   27
      Top             =   540
      Visible         =   0   'False
      Width           =   105
      _cx             =   185
      _cy             =   14764
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   200
      ColWidthMax     =   200
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
   Begin VB.Timer ScrolTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   990
      Top             =   0
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1285
      Index           =   4
      Left            =   6030
      TabIndex        =   22
      Top             =   2340
      Width           =   1860
      _Version        =   524298
      _ExtentX        =   3281
      _ExtentY        =   2267
      _StockProps     =   66
      Caption         =   "Close Table"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":10351
      textLT          =   "frmInput.frx":103C7
      textCT          =   "frmInput.frx":103DF
      textRT          =   "frmInput.frx":103F7
      textLM          =   "frmInput.frx":1040F
      textRM          =   "frmInput.frx":10427
      textLB          =   "frmInput.frx":1043F
      textCB          =   "frmInput.frx":10457
      textRB          =   "frmInput.frx":1046F
      colorBack       =   "frmInput.frx":10487
      colorIntern     =   "frmInput.frx":104B1
      colorMO         =   "frmInput.frx":104DB
      colorFocus      =   "frmInput.frx":10505
      colorDisabled   =   "frmInput.frx":1052F
      colorPressed    =   "frmInput.frx":10559
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.Timer errTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3720
      Top             =   90
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1950
      Top             =   60
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1425
      Index           =   11
      Left            =   13410
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9075
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   2514
      _StockProps     =   66
      Caption         =   "Exit"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      CornerFactor    =   15
      BackColorContainer=   11983853
      SpecialEffect   =   1
      CaptionWordWrapPerc=   95
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":10583
      textLT          =   "frmInput.frx":105EB
      textCT          =   "frmInput.frx":10603
      textRT          =   "frmInput.frx":1061B
      textLM          =   "frmInput.frx":10633
      textRM          =   "frmInput.frx":1064B
      textLB          =   "frmInput.frx":10663
      textCB          =   "frmInput.frx":1067B
      textRB          =   "frmInput.frx":10693
      colorBack       =   "frmInput.frx":106AB
      colorIntern     =   "frmInput.frx":106D5
      colorMO         =   "frmInput.frx":106FF
      colorFocus      =   "frmInput.frx":10729
      colorDisabled   =   "frmInput.frx":10753
      colorPressed    =   "frmInput.frx":1077D
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   8
      Left            =   13410
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7710
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "9"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":107A7
      textLT          =   "frmInput.frx":10809
      textCT          =   "frmInput.frx":10821
      textRT          =   "frmInput.frx":10839
      textLM          =   "frmInput.frx":10851
      textRM          =   "frmInput.frx":10869
      textLB          =   "frmInput.frx":10881
      textCB          =   "frmInput.frx":10899
      textRB          =   "frmInput.frx":108B1
      colorBack       =   "frmInput.frx":108C9
      colorIntern     =   "frmInput.frx":108F3
      colorMO         =   "frmInput.frx":1091D
      colorFocus      =   "frmInput.frx":10947
      colorDisabled   =   "frmInput.frx":10971
      colorPressed    =   "frmInput.frx":1099B
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   2
      Left            =   13410
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4980
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "3"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      CornerFactor    =   15
      BackColorContainer=   11983853
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":109C5
      textLT          =   "frmInput.frx":10A27
      textCT          =   "frmInput.frx":10A3F
      textRT          =   "frmInput.frx":10A57
      textLM          =   "frmInput.frx":10A6F
      textRM          =   "frmInput.frx":10A87
      textLB          =   "frmInput.frx":10A9F
      textCB          =   "frmInput.frx":10AB7
      textRB          =   "frmInput.frx":10ACF
      colorBack       =   "frmInput.frx":10AE7
      colorIntern     =   "frmInput.frx":10B11
      colorMO         =   "frmInput.frx":10B3B
      colorFocus      =   "frmInput.frx":10B65
      colorDisabled   =   "frmInput.frx":10B8F
      colorPressed    =   "frmInput.frx":10BB9
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   5
      Left            =   13410
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6345
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "6"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":10BE3
      textLT          =   "frmInput.frx":10C45
      textCT          =   "frmInput.frx":10C5D
      textRT          =   "frmInput.frx":10C75
      textLM          =   "frmInput.frx":10C8D
      textRM          =   "frmInput.frx":10CA5
      textLB          =   "frmInput.frx":10CBD
      textCB          =   "frmInput.frx":10CD5
      textRB          =   "frmInput.frx":10CED
      colorBack       =   "frmInput.frx":10D05
      colorIntern     =   "frmInput.frx":10D2F
      colorMO         =   "frmInput.frx":10D59
      colorFocus      =   "frmInput.frx":10D83
      colorDisabled   =   "frmInput.frx":10DAD
      colorPressed    =   "frmInput.frx":10DD7
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1425
      Index           =   10
      Left            =   11715
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9075
      Width           =   1695
      _Version        =   524298
      _ExtentX        =   2990
      _ExtentY        =   2514
      _StockProps     =   66
      Caption         =   "0"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":10E01
      textLT          =   "frmInput.frx":10E63
      textCT          =   "frmInput.frx":10E7B
      textRT          =   "frmInput.frx":10E93
      textLM          =   "frmInput.frx":10EAB
      textRM          =   "frmInput.frx":10EC3
      textLB          =   "frmInput.frx":10EDB
      textCB          =   "frmInput.frx":10EF3
      textRB          =   "frmInput.frx":10F0B
      colorBack       =   "frmInput.frx":10F23
      colorIntern     =   "frmInput.frx":10F4D
      colorMO         =   "frmInput.frx":10F77
      colorFocus      =   "frmInput.frx":10FA1
      colorDisabled   =   "frmInput.frx":10FCB
      colorPressed    =   "frmInput.frx":10FF5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   7
      Left            =   11715
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7710
      Width           =   1695
      _Version        =   524298
      _ExtentX        =   2990
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "8"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":1101F
      textLT          =   "frmInput.frx":11081
      textCT          =   "frmInput.frx":11099
      textRT          =   "frmInput.frx":110B1
      textLM          =   "frmInput.frx":110C9
      textRM          =   "frmInput.frx":110E1
      textLB          =   "frmInput.frx":110F9
      textCB          =   "frmInput.frx":11111
      textRB          =   "frmInput.frx":11129
      colorBack       =   "frmInput.frx":11141
      colorIntern     =   "frmInput.frx":1116B
      colorMO         =   "frmInput.frx":11195
      colorFocus      =   "frmInput.frx":111BF
      colorDisabled   =   "frmInput.frx":111E9
      colorPressed    =   "frmInput.frx":11213
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   1
      Left            =   11715
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4980
      Width           =   1695
      _Version        =   524298
      _ExtentX        =   2990
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "2"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":1123D
      textLT          =   "frmInput.frx":1129F
      textCT          =   "frmInput.frx":112B7
      textRT          =   "frmInput.frx":112CF
      textLM          =   "frmInput.frx":112E7
      textRM          =   "frmInput.frx":112FF
      textLB          =   "frmInput.frx":11317
      textCB          =   "frmInput.frx":1132F
      textRB          =   "frmInput.frx":11347
      colorBack       =   "frmInput.frx":1135F
      colorIntern     =   "frmInput.frx":11389
      colorMO         =   "frmInput.frx":113B3
      colorFocus      =   "frmInput.frx":113DD
      colorDisabled   =   "frmInput.frx":11407
      colorPressed    =   "frmInput.frx":11431
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   4
      Left            =   11715
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6345
      Width           =   1695
      _Version        =   524298
      _ExtentX        =   2990
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "5"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":1145B
      textLT          =   "frmInput.frx":114BD
      textCT          =   "frmInput.frx":114D5
      textRT          =   "frmInput.frx":114ED
      textLM          =   "frmInput.frx":11505
      textRM          =   "frmInput.frx":1151D
      textLB          =   "frmInput.frx":11535
      textCB          =   "frmInput.frx":1154D
      textRB          =   "frmInput.frx":11565
      colorBack       =   "frmInput.frx":1157D
      colorIntern     =   "frmInput.frx":115A7
      colorMO         =   "frmInput.frx":115D1
      colorFocus      =   "frmInput.frx":115FB
      colorDisabled   =   "frmInput.frx":11625
      colorPressed    =   "frmInput.frx":1164F
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1425
      Index           =   9
      Left            =   9990
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9075
      Width           =   1725
      _Version        =   524298
      _ExtentX        =   3043
      _ExtentY        =   2514
      _StockProps     =   66
      Caption         =   "CL"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      CornerFactor    =   15
      BackColorContainer=   11983853
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":11679
      textLT          =   "frmInput.frx":116DD
      textCT          =   "frmInput.frx":116F5
      textRT          =   "frmInput.frx":1170D
      textLM          =   "frmInput.frx":11725
      textRM          =   "frmInput.frx":1173D
      textLB          =   "frmInput.frx":11755
      textCB          =   "frmInput.frx":1176D
      textRB          =   "frmInput.frx":11785
      colorBack       =   "frmInput.frx":1179D
      colorIntern     =   "frmInput.frx":117C7
      colorMO         =   "frmInput.frx":117F1
      colorFocus      =   "frmInput.frx":1181B
      colorDisabled   =   "frmInput.frx":11845
      colorPressed    =   "frmInput.frx":1186F
      Orientation     =   8
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   6
      Left            =   9990
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7710
      Width           =   1725
      _Version        =   524298
      _ExtentX        =   3043
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "7"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":11899
      textLT          =   "frmInput.frx":118FB
      textCT          =   "frmInput.frx":11913
      textRT          =   "frmInput.frx":1192B
      textLM          =   "frmInput.frx":11943
      textRM          =   "frmInput.frx":1195B
      textLB          =   "frmInput.frx":11973
      textCB          =   "frmInput.frx":1198B
      textRB          =   "frmInput.frx":119A3
      colorBack       =   "frmInput.frx":119BB
      colorIntern     =   "frmInput.frx":119E5
      colorMO         =   "frmInput.frx":11A0F
      colorFocus      =   "frmInput.frx":11A39
      colorDisabled   =   "frmInput.frx":11A63
      colorPressed    =   "frmInput.frx":11A8D
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   3
      Left            =   9990
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6345
      Width           =   1725
      _Version        =   524298
      _ExtentX        =   3043
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "4"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorContainer=   16777215
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":11AB7
      textLT          =   "frmInput.frx":11B19
      textCT          =   "frmInput.frx":11B31
      textRT          =   "frmInput.frx":11B49
      textLM          =   "frmInput.frx":11B61
      textRM          =   "frmInput.frx":11B79
      textLB          =   "frmInput.frx":11B91
      textCB          =   "frmInput.frx":11BA9
      textRB          =   "frmInput.frx":11BC1
      colorBack       =   "frmInput.frx":11BD9
      colorIntern     =   "frmInput.frx":11C03
      colorMO         =   "frmInput.frx":11C2D
      colorFocus      =   "frmInput.frx":11C57
      colorDisabled   =   "frmInput.frx":11C81
      colorPressed    =   "frmInput.frx":11CAB
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1365
      Index           =   0
      Left            =   9990
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4980
      Width           =   1725
      _Version        =   524298
      _ExtentX        =   3043
      _ExtentY        =   2408
      _StockProps     =   66
      Caption         =   "1"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   11983853
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":11CD5
      textLT          =   "frmInput.frx":11D37
      textCT          =   "frmInput.frx":11D4F
      textRT          =   "frmInput.frx":11D67
      textLM          =   "frmInput.frx":11D7F
      textRM          =   "frmInput.frx":11D97
      textLB          =   "frmInput.frx":11DAF
      textCB          =   "frmInput.frx":11DC7
      textRB          =   "frmInput.frx":11DDF
      colorBack       =   "frmInput.frx":11DF7
      colorIntern     =   "frmInput.frx":11E21
      colorMO         =   "frmInput.frx":11E4B
      colorFocus      =   "frmInput.frx":11E75
      colorDisabled   =   "frmInput.frx":11E9F
      colorPressed    =   "frmInput.frx":11EC9
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   975
      Index           =   12
      Left            =   9990
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1500
      Width           =   5025
      _Version        =   524298
      _ExtentX        =   8864
      _ExtentY        =   1720
      _StockProps     =   66
      Caption         =   "Reservations"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shape           =   1
      CornerFactor    =   38
      Surface         =   1
      BackColorContainer=   13756915
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":11EF3
      textLT          =   "frmInput.frx":11F6B
      textCT          =   "frmInput.frx":11F83
      textRT          =   "frmInput.frx":11F9B
      textLM          =   "frmInput.frx":11FB3
      textRM          =   "frmInput.frx":11FCB
      textLB          =   "frmInput.frx":11FE3
      textCB          =   "frmInput.frx":11FFB
      textRB          =   "frmInput.frx":12013
      colorBack       =   "frmInput.frx":1202B
      colorIntern     =   "frmInput.frx":12055
      colorMO         =   "frmInput.frx":1207F
      colorFocus      =   "frmInput.frx":120A9
      colorDisabled   =   "frmInput.frx":120D3
      colorPressed    =   "frmInput.frx":120FD
      HollowFrame     =   -1  'True
      LightDirection  =   1
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   0
         Top             =   180
      End
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1285
      Index           =   1
      Left            =   390
      TabIndex        =   18
      Top             =   2340
      Width           =   1920
      _Version        =   524298
      _ExtentX        =   3387
      _ExtentY        =   2267
      _StockProps     =   66
      Caption         =   "Show All Tables"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":12127
      textLT          =   "frmInput.frx":121A5
      textCT          =   "frmInput.frx":121BD
      textRT          =   "frmInput.frx":121D5
      textLM          =   "frmInput.frx":121ED
      textRM          =   "frmInput.frx":12205
      textLB          =   "frmInput.frx":1221D
      textCB          =   "frmInput.frx":12235
      textRB          =   "frmInput.frx":1224D
      colorBack       =   "frmInput.frx":12265
      colorIntern     =   "frmInput.frx":1228F
      colorMO         =   "frmInput.frx":122B9
      colorFocus      =   "frmInput.frx":122E3
      colorDisabled   =   "frmInput.frx":1230D
      colorPressed    =   "frmInput.frx":12337
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1285
      Index           =   5
      Left            =   4170
      TabIndex        =   23
      Top             =   2340
      Width           =   1860
      _Version        =   524298
      _ExtentX        =   3281
      _ExtentY        =   2267
      _StockProps     =   66
      Caption         =   "Split Bill"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   19
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":12361
      textLT          =   "frmInput.frx":123D5
      textCT          =   "frmInput.frx":123ED
      textRT          =   "frmInput.frx":12405
      textLM          =   "frmInput.frx":1241D
      textRM          =   "frmInput.frx":12435
      textLB          =   "frmInput.frx":1244D
      textCB          =   "frmInput.frx":12465
      textRB          =   "frmInput.frx":1247D
      colorBack       =   "frmInput.frx":12495
      colorIntern     =   "frmInput.frx":124BF
      colorMO         =   "frmInput.frx":124E9
      colorFocus      =   "frmInput.frx":12513
      colorDisabled   =   "frmInput.frx":1253D
      colorPressed    =   "frmInput.frx":12567
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.PictureBox picHoldFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3EEEF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1950
      ScaleHeight     =   615
      ScaleWidth      =   825
      TabIndex        =   19
      Top             =   1140
      Width           =   825
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1290
      Index           =   7
      Left            =   7890
      TabIndex        =   25
      Top             =   2340
      Width           =   1740
      _Version        =   524298
      _ExtentX        =   3069
      _ExtentY        =   2275
      _StockProps     =   66
      Caption         =   "View Table"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   19
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":12591
      textLT          =   "frmInput.frx":12605
      textCT          =   "frmInput.frx":1261D
      textRT          =   "frmInput.frx":12635
      textLM          =   "frmInput.frx":1264D
      textRM          =   "frmInput.frx":12665
      textLB          =   "frmInput.frx":1267D
      textCB          =   "frmInput.frx":12695
      textRB          =   "frmInput.frx":126AD
      colorBack       =   "frmInput.frx":126C5
      colorIntern     =   "frmInput.frx":126EF
      colorMO         =   "frmInput.frx":12719
      colorFocus      =   "frmInput.frx":12743
      colorDisabled   =   "frmInput.frx":1276D
      colorPressed    =   "frmInput.frx":12797
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   4
      Left            =   360
      TabIndex        =   28
      Top             =   5085
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":127C1
      textLT          =   "frmInput.frx":127D9
      textCT          =   "frmInput.frx":127F1
      textRT          =   "frmInput.frx":12809
      textLM          =   "frmInput.frx":12821
      textRM          =   "frmInput.frx":12839
      textLB          =   "frmInput.frx":12851
      textCB          =   "frmInput.frx":12869
      textRB          =   "frmInput.frx":12881
      colorBack       =   "frmInput.frx":12899
      colorIntern     =   "frmInput.frx":128C3
      colorMO         =   "frmInput.frx":128ED
      colorFocus      =   "frmInput.frx":12917
      colorDisabled   =   "frmInput.frx":12941
      colorPressed    =   "frmInput.frx":1296B
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   5
      Left            =   2730
      TabIndex        =   29
      Top             =   5085
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":12995
      textLT          =   "frmInput.frx":129AD
      textCT          =   "frmInput.frx":129C5
      textRT          =   "frmInput.frx":129DD
      textLM          =   "frmInput.frx":129F5
      textRM          =   "frmInput.frx":12A0D
      textLB          =   "frmInput.frx":12A25
      textCB          =   "frmInput.frx":12A3D
      textRB          =   "frmInput.frx":12A55
      colorBack       =   "frmInput.frx":12A6D
      colorIntern     =   "frmInput.frx":12A97
      colorMO         =   "frmInput.frx":12AC1
      colorFocus      =   "frmInput.frx":12AEB
      colorDisabled   =   "frmInput.frx":12B15
      colorPressed    =   "frmInput.frx":12B3F
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   6
      Left            =   5070
      TabIndex        =   30
      Top             =   5085
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":12B69
      textLT          =   "frmInput.frx":12B81
      textCT          =   "frmInput.frx":12B99
      textRT          =   "frmInput.frx":12BB1
      textLM          =   "frmInput.frx":12BC9
      textRM          =   "frmInput.frx":12BE1
      textLB          =   "frmInput.frx":12BF9
      textCB          =   "frmInput.frx":12C11
      textRB          =   "frmInput.frx":12C29
      colorBack       =   "frmInput.frx":12C41
      colorIntern     =   "frmInput.frx":12C6B
      colorMO         =   "frmInput.frx":12C95
      colorFocus      =   "frmInput.frx":12CBF
      colorDisabled   =   "frmInput.frx":12CE9
      colorPressed    =   "frmInput.frx":12D13
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   7
      Left            =   7440
      TabIndex        =   31
      Top             =   5085
      Visible         =   0   'False
      Width           =   2310
      _Version        =   524298
      _ExtentX        =   4075
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":12D3D
      textLT          =   "frmInput.frx":12D55
      textCT          =   "frmInput.frx":12D6D
      textRT          =   "frmInput.frx":12D85
      textLM          =   "frmInput.frx":12D9D
      textRM          =   "frmInput.frx":12DB5
      textLB          =   "frmInput.frx":12DCD
      textCB          =   "frmInput.frx":12DE5
      textRB          =   "frmInput.frx":12DFD
      colorBack       =   "frmInput.frx":12E15
      colorIntern     =   "frmInput.frx":12E3F
      colorMO         =   "frmInput.frx":12E69
      colorFocus      =   "frmInput.frx":12E93
      colorDisabled   =   "frmInput.frx":12EBD
      colorPressed    =   "frmInput.frx":12EE7
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   8
      Left            =   360
      TabIndex        =   32
      Top             =   6465
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":12F11
      textLT          =   "frmInput.frx":12F29
      textCT          =   "frmInput.frx":12F41
      textRT          =   "frmInput.frx":12F59
      textLM          =   "frmInput.frx":12F71
      textRM          =   "frmInput.frx":12F89
      textLB          =   "frmInput.frx":12FA1
      textCB          =   "frmInput.frx":12FB9
      textRB          =   "frmInput.frx":12FD1
      colorBack       =   "frmInput.frx":12FE9
      colorIntern     =   "frmInput.frx":13013
      colorMO         =   "frmInput.frx":1303D
      colorFocus      =   "frmInput.frx":13067
      colorDisabled   =   "frmInput.frx":13091
      colorPressed    =   "frmInput.frx":130BB
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   9
      Left            =   2730
      TabIndex        =   33
      Top             =   6465
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":130E5
      textLT          =   "frmInput.frx":130FD
      textCT          =   "frmInput.frx":13115
      textRT          =   "frmInput.frx":1312D
      textLM          =   "frmInput.frx":13145
      textRM          =   "frmInput.frx":1315D
      textLB          =   "frmInput.frx":13175
      textCB          =   "frmInput.frx":1318D
      textRB          =   "frmInput.frx":131A5
      colorBack       =   "frmInput.frx":131BD
      colorIntern     =   "frmInput.frx":131E7
      colorMO         =   "frmInput.frx":13211
      colorFocus      =   "frmInput.frx":1323B
      colorDisabled   =   "frmInput.frx":13265
      colorPressed    =   "frmInput.frx":1328F
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   10
      Left            =   5070
      TabIndex        =   34
      Top             =   6465
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":132B9
      textLT          =   "frmInput.frx":132D1
      textCT          =   "frmInput.frx":132E9
      textRT          =   "frmInput.frx":13301
      textLM          =   "frmInput.frx":13319
      textRM          =   "frmInput.frx":13331
      textLB          =   "frmInput.frx":13349
      textCB          =   "frmInput.frx":13361
      textRB          =   "frmInput.frx":13379
      colorBack       =   "frmInput.frx":13391
      colorIntern     =   "frmInput.frx":133BB
      colorMO         =   "frmInput.frx":133E5
      colorFocus      =   "frmInput.frx":1340F
      colorDisabled   =   "frmInput.frx":13439
      colorPressed    =   "frmInput.frx":13463
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   11
      Left            =   7440
      TabIndex        =   35
      Top             =   6465
      Visible         =   0   'False
      Width           =   2310
      _Version        =   524298
      _ExtentX        =   4075
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":1348D
      textLT          =   "frmInput.frx":134A5
      textCT          =   "frmInput.frx":134BD
      textRT          =   "frmInput.frx":134D5
      textLM          =   "frmInput.frx":134ED
      textRM          =   "frmInput.frx":13505
      textLB          =   "frmInput.frx":1351D
      textCB          =   "frmInput.frx":13535
      textRB          =   "frmInput.frx":1354D
      colorBack       =   "frmInput.frx":13565
      colorIntern     =   "frmInput.frx":1358F
      colorMO         =   "frmInput.frx":135B9
      colorFocus      =   "frmInput.frx":135E3
      colorDisabled   =   "frmInput.frx":1360D
      colorPressed    =   "frmInput.frx":13637
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   12
      Left            =   360
      TabIndex        =   36
      Top             =   7845
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":13661
      textLT          =   "frmInput.frx":13679
      textCT          =   "frmInput.frx":13691
      textRT          =   "frmInput.frx":136A9
      textLM          =   "frmInput.frx":136C1
      textRM          =   "frmInput.frx":136D9
      textLB          =   "frmInput.frx":136F1
      textCB          =   "frmInput.frx":13709
      textRB          =   "frmInput.frx":13721
      colorBack       =   "frmInput.frx":13739
      colorIntern     =   "frmInput.frx":13763
      colorMO         =   "frmInput.frx":1378D
      colorFocus      =   "frmInput.frx":137B7
      colorDisabled   =   "frmInput.frx":137E1
      colorPressed    =   "frmInput.frx":1380B
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   13
      Left            =   2730
      TabIndex        =   37
      Top             =   7845
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":13835
      textLT          =   "frmInput.frx":1384D
      textCT          =   "frmInput.frx":13865
      textRT          =   "frmInput.frx":1387D
      textLM          =   "frmInput.frx":13895
      textRM          =   "frmInput.frx":138AD
      textLB          =   "frmInput.frx":138C5
      textCB          =   "frmInput.frx":138DD
      textRB          =   "frmInput.frx":138F5
      colorBack       =   "frmInput.frx":1390D
      colorIntern     =   "frmInput.frx":13937
      colorMO         =   "frmInput.frx":13961
      colorFocus      =   "frmInput.frx":1398B
      colorDisabled   =   "frmInput.frx":139B5
      colorPressed    =   "frmInput.frx":139DF
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   14
      Left            =   5070
      TabIndex        =   38
      Top             =   7845
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":13A09
      textLT          =   "frmInput.frx":13A21
      textCT          =   "frmInput.frx":13A39
      textRT          =   "frmInput.frx":13A51
      textLM          =   "frmInput.frx":13A69
      textRM          =   "frmInput.frx":13A81
      textLB          =   "frmInput.frx":13A99
      textCB          =   "frmInput.frx":13AB1
      textRB          =   "frmInput.frx":13AC9
      colorBack       =   "frmInput.frx":13AE1
      colorIntern     =   "frmInput.frx":13B0B
      colorMO         =   "frmInput.frx":13B35
      colorFocus      =   "frmInput.frx":13B5F
      colorDisabled   =   "frmInput.frx":13B89
      colorPressed    =   "frmInput.frx":13BB3
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   15
      Left            =   7440
      TabIndex        =   39
      Top             =   7845
      Visible         =   0   'False
      Width           =   2310
      _Version        =   524298
      _ExtentX        =   4075
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":13BDD
      textLT          =   "frmInput.frx":13BF5
      textCT          =   "frmInput.frx":13C0D
      textRT          =   "frmInput.frx":13C25
      textLM          =   "frmInput.frx":13C3D
      textRM          =   "frmInput.frx":13C55
      textLB          =   "frmInput.frx":13C6D
      textCB          =   "frmInput.frx":13C85
      textRB          =   "frmInput.frx":13C9D
      colorBack       =   "frmInput.frx":13CB5
      colorIntern     =   "frmInput.frx":13CDF
      colorMO         =   "frmInput.frx":13D09
      colorFocus      =   "frmInput.frx":13D33
      colorDisabled   =   "frmInput.frx":13D5D
      colorPressed    =   "frmInput.frx":13D87
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   16
      Left            =   360
      TabIndex        =   40
      Top             =   9180
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":13DB1
      textLT          =   "frmInput.frx":13DC9
      textCT          =   "frmInput.frx":13DE1
      textRT          =   "frmInput.frx":13DF9
      textLM          =   "frmInput.frx":13E11
      textRM          =   "frmInput.frx":13E29
      textLB          =   "frmInput.frx":13E41
      textCB          =   "frmInput.frx":13E59
      textRB          =   "frmInput.frx":13E71
      colorBack       =   "frmInput.frx":13E89
      colorIntern     =   "frmInput.frx":13EB3
      colorMO         =   "frmInput.frx":13EDD
      colorFocus      =   "frmInput.frx":13F07
      colorDisabled   =   "frmInput.frx":13F31
      colorPressed    =   "frmInput.frx":13F5B
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   17
      Left            =   2730
      TabIndex        =   41
      Top             =   9180
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":13F85
      textLT          =   "frmInput.frx":13F9D
      textCT          =   "frmInput.frx":13FB5
      textRT          =   "frmInput.frx":13FCD
      textLM          =   "frmInput.frx":13FE5
      textRM          =   "frmInput.frx":13FFD
      textLB          =   "frmInput.frx":14015
      textCB          =   "frmInput.frx":1402D
      textRB          =   "frmInput.frx":14045
      colorBack       =   "frmInput.frx":1405D
      colorIntern     =   "frmInput.frx":14087
      colorMO         =   "frmInput.frx":140B1
      colorFocus      =   "frmInput.frx":140DB
      colorDisabled   =   "frmInput.frx":14105
      colorPressed    =   "frmInput.frx":1412F
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   18
      Left            =   5070
      TabIndex        =   42
      Top             =   9180
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":14159
      textLT          =   "frmInput.frx":14171
      textCT          =   "frmInput.frx":14189
      textRT          =   "frmInput.frx":141A1
      textLM          =   "frmInput.frx":141B9
      textRM          =   "frmInput.frx":141D1
      textLB          =   "frmInput.frx":141E9
      textCB          =   "frmInput.frx":14201
      textRB          =   "frmInput.frx":14219
      colorBack       =   "frmInput.frx":14231
      colorIntern     =   "frmInput.frx":1425B
      colorMO         =   "frmInput.frx":14285
      colorFocus      =   "frmInput.frx":142AF
      colorDisabled   =   "frmInput.frx":142D9
      colorPressed    =   "frmInput.frx":14303
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   19
      Left            =   7440
      TabIndex        =   43
      Top             =   9180
      Visible         =   0   'False
      Width           =   2310
      _Version        =   524298
      _ExtentX        =   4075
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":1432D
      textLT          =   "frmInput.frx":14345
      textCT          =   "frmInput.frx":1435D
      textRT          =   "frmInput.frx":14375
      textLM          =   "frmInput.frx":1438D
      textRM          =   "frmInput.frx":143A5
      textLB          =   "frmInput.frx":143BD
      textCB          =   "frmInput.frx":143D5
      textRB          =   "frmInput.frx":143ED
      colorBack       =   "frmInput.frx":14405
      colorIntern     =   "frmInput.frx":1442F
      colorMO         =   "frmInput.frx":14459
      colorFocus      =   "frmInput.frx":14483
      colorDisabled   =   "frmInput.frx":144AD
      colorPressed    =   "frmInput.frx":144D7
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   1
      Left            =   2730
      TabIndex        =   44
      Top             =   3735
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":14501
      textLT          =   "frmInput.frx":14519
      textCT          =   "frmInput.frx":14531
      textRT          =   "frmInput.frx":14549
      textLM          =   "frmInput.frx":14561
      textRM          =   "frmInput.frx":14579
      textLB          =   "frmInput.frx":14591
      textCB          =   "frmInput.frx":145A9
      textRB          =   "frmInput.frx":145C1
      colorBack       =   "frmInput.frx":145D9
      colorIntern     =   "frmInput.frx":14603
      colorMO         =   "frmInput.frx":1462D
      colorFocus      =   "frmInput.frx":14657
      colorDisabled   =   "frmInput.frx":14681
      colorPressed    =   "frmInput.frx":146AB
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   2
      Left            =   5070
      TabIndex        =   45
      Top             =   3735
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":146D5
      textLT          =   "frmInput.frx":146ED
      textCT          =   "frmInput.frx":14705
      textRT          =   "frmInput.frx":1471D
      textLM          =   "frmInput.frx":14735
      textRM          =   "frmInput.frx":1474D
      textLB          =   "frmInput.frx":14765
      textCB          =   "frmInput.frx":1477D
      textRB          =   "frmInput.frx":14795
      colorBack       =   "frmInput.frx":147AD
      colorIntern     =   "frmInput.frx":147D7
      colorMO         =   "frmInput.frx":14801
      colorFocus      =   "frmInput.frx":1482B
      colorDisabled   =   "frmInput.frx":14855
      colorPressed    =   "frmInput.frx":1487F
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   0
      Left            =   360
      TabIndex        =   46
      Top             =   3735
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":148A9
      textLT          =   "frmInput.frx":148C1
      textCT          =   "frmInput.frx":148D9
      textRT          =   "frmInput.frx":148F1
      textLM          =   "frmInput.frx":14909
      textRM          =   "frmInput.frx":14921
      textLB          =   "frmInput.frx":14939
      textCB          =   "frmInput.frx":14951
      textRB          =   "frmInput.frx":14969
      colorBack       =   "frmInput.frx":14981
      colorIntern     =   "frmInput.frx":149AB
      colorMO         =   "frmInput.frx":149D5
      colorFocus      =   "frmInput.frx":149FF
      colorDisabled   =   "frmInput.frx":14A29
      colorPressed    =   "frmInput.frx":14A53
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTable 
      Height          =   1380
      Index           =   3
      Left            =   7440
      TabIndex        =   47
      Top             =   3735
      Visible         =   0   'False
      Width           =   2310
      _Version        =   524298
      _ExtentX        =   4075
      _ExtentY        =   2434
      _StockProps     =   66
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":14A7D
      textLT          =   "frmInput.frx":14A95
      textCT          =   "frmInput.frx":14AAD
      textRT          =   "frmInput.frx":14AC5
      textLM          =   "frmInput.frx":14ADD
      textRM          =   "frmInput.frx":14AF5
      textLB          =   "frmInput.frx":14B0D
      textCB          =   "frmInput.frx":14B25
      textRB          =   "frmInput.frx":14B3D
      colorBack       =   "frmInput.frx":14B55
      colorIntern     =   "frmInput.frx":14B7F
      colorMO         =   "frmInput.frx":14BA9
      colorFocus      =   "frmInput.frx":14BD3
      colorDisabled   =   "frmInput.frx":14BFD
      colorPressed    =   "frmInput.frx":14C27
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1285
      Index           =   0
      Left            =   2310
      TabIndex        =   59
      Top             =   2340
      Width           =   1860
      _Version        =   524298
      _ExtentX        =   3281
      _ExtentY        =   2267
      _StockProps     =   66
      Caption         =   "Change Waiter"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   18
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":14C51
      textLT          =   "frmInput.frx":14CCB
      textCT          =   "frmInput.frx":14CE3
      textRT          =   "frmInput.frx":14CFB
      textLM          =   "frmInput.frx":14D13
      textRM          =   "frmInput.frx":14D2B
      textLB          =   "frmInput.frx":14D43
      textCB          =   "frmInput.frx":14D5B
      textRB          =   "frmInput.frx":14D73
      colorBack       =   "frmInput.frx":14D8B
      colorIntern     =   "frmInput.frx":14DB5
      colorMO         =   "frmInput.frx":14DDF
      colorFocus      =   "frmInput.frx":14E09
      colorDisabled   =   "frmInput.frx":14E33
      colorPressed    =   "frmInput.frx":14E5D
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   945
      Index           =   2
      Left            =   8040
      TabIndex        =   65
      Top             =   1440
      Width           =   1725
      _Version        =   524298
      _ExtentX        =   3043
      _ExtentY        =   1667
      _StockProps     =   66
      Caption         =   "Name Table"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRM {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextLB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextCB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontTextRB {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CornerFactor    =   20
      Surface         =   1
      BackColorContainer=   14737632
      SpecialEffect   =   1
      CaptionWordWrapPerc=   100
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput.frx":14E87
      textLT          =   "frmInput.frx":14EFB
      textCT          =   "frmInput.frx":14F13
      textRT          =   "frmInput.frx":14F2B
      textLM          =   "frmInput.frx":14F43
      textRM          =   "frmInput.frx":14F5B
      textLB          =   "frmInput.frx":14F73
      textCB          =   "frmInput.frx":14F8B
      textRB          =   "frmInput.frx":14FA3
      colorBack       =   "frmInput.frx":14FBB
      colorIntern     =   "frmInput.frx":14FE5
      colorMO         =   "frmInput.frx":1500F
      colorFocus      =   "frmInput.frx":15039
      colorDisabled   =   "frmInput.frx":15063
      colorPressed    =   "frmInput.frx":1508D
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin MSForms.Label lblTables 
      Height          =   315
      Left            =   4920
      TabIndex        =   63
      Top             =   10830
      Width           =   5565
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "9816;556"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label lblUser 
      Height          =   315
      Left            =   10590
      TabIndex        =   62
      Top             =   10830
      Width           =   4365
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7699;556"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblDate 
      Height          =   315
      Left            =   450
      TabIndex        =   61
      Top             =   10830
      Width           =   4365
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7699;556"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblKeyRegister 
      Height          =   645
      Left            =   4860
      TabIndex        =   49
      Top             =   480
      Width           =   10005
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "17648;1138"
      FontName        =   "Arial Narrow"
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image4 
      Height          =   1185
      Left            =   11610
      Top             =   3690
      Width           =   1545
      BorderStyle     =   0
      Size            =   "2725;2090"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image3 
      Height          =   1185
      Left            =   11610
      Top             =   2520
      Width           =   1545
      BorderStyle     =   0
      Size            =   "2725;2090"
      VariousPropertyBits=   19
   End
   Begin MSForms.Label Label2 
      Height          =   435
      Left            =   9990
      TabIndex        =   15
      Top             =   4110
      Width           =   1575
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Covers:"
      Size            =   "2778;767"
      FontName        =   "Arial Narrow"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   435
      Left            =   9990
      TabIndex        =   14
      Top             =   3000
      Width           =   1575
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Table No:"
      Size            =   "2778;767"
      FontName        =   "Arial Narrow"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image2 
      Height          =   1125
      Left            =   11640
      Top             =   3750
      Width           =   1545
      BackColor       =   16777215
      Size            =   "2725;1984"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image1 
      Height          =   1125
      Left            =   11640
      Top             =   2580
      Width           =   1545
      BorderColor     =   255
      BackColor       =   16777215
      Size            =   "2725;1984"
      VariousPropertyBits=   19
   End
   Begin MSForms.TextBox txtCovers 
      Height          =   705
      Left            =   11940
      TabIndex        =   2
      Top             =   4080
      Width           =   1245
      VariousPropertyBits=   746604567
      MaxLength       =   4
      Size            =   "2196;1244"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   405
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtTables 
      Height          =   705
      Left            =   11940
      TabIndex        =   0
      Top             =   2940
      Width           =   1245
      VariousPropertyBits=   746604567
      MaxLength       =   4
      Size            =   "2196;1244"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   405
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblServers 
      Caption         =   "Label3"
      Height          =   585
      Left            =   180
      TabIndex        =   58
      Top             =   2970
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape shpLive 
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   2
      Left            =   1170
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin VB.Shape shpLive 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   1
      Left            =   915
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin VB.Shape shpLive 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   225
      Index           =   0
      Left            =   660
      Shape           =   3  'Circle
      Top             =   300
      Width           =   165
   End
   Begin MSForms.Image newBack 
      Height          =   1395
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "1984;2461"
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdArrow_Click(Index As Integer)
    If grdMain.Rows = 1 Then Exit Sub
    If grdMain.Enabled = False Then
        voidTimer.Enabled = False
        grdMain.Enabled = True
        grdMain.Cell(flexcpBackColor, grdMain.Row, 0, grdMain.Row, 2) = 0
        grdMain.HighLight = flexHighlightAlways
    End If
    
    grdMain.SetFocus
    
    If grdMain.Rows > 2 And grdMain.Row = 0 Then
        grdMain.Row = 1
    End If
    Select Case Index
        Case 0
            If grdMain.Row > 1 Then
                grdMain.Row = grdMain.Row - 1
            End If
            grdMain.ShowCell grdMain.Row, 0
        Case 1
            If grdMain.Row < grdMain.Rows - 1 Then
                grdMain.Row = grdMain.Row + 1
            End If
            grdMain.ShowCell grdMain.Row - 1, 0
    End Select
End Sub
Private Sub cmdClose_Click()
    picSlip.Visible = False
    ActiveReadServer "Select * from Table_Listing_View where User_No = " & UserRecord.User_Number
    lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Tables"
    rs.Close
End Sub
Private Sub cmdErr_Click()
    cmdErr.Visible = False
    cmdErr.BackColor = &HF2&
    errTimer.Enabled = False
    cmdErr.Caption = ""
    If txtTables.Enabled = True Then txtTables.SetFocus
    If cmdFancy(3).Enabled = False Then
        errTimer.Enabled = False
        cmdFancy(1).Enabled = True
        cmdFancy(3).Enabled = True
        cmdFancy(4).Enabled = True
        cmdFancy(5).Enabled = True
        cmdFancy(6).Enabled = True
        cmdFancy(7).Enabled = True
        cmdFancy(8).Enabled = True
        txtCovers.Enabled = True
        cmdLogoff.Orientation = DIR_NW
        Select Case cmdFancy(1).Caption
            Case "Show All Tables"
                LoadTables 0
            Case "Show Own Tables"
                LoadTables 1
        End Select
        lblTransfer.Tag = ""
    End If
End Sub
Private Sub cmdFancy_Click(Index As Integer)
    If txtTables.Enabled = False Then
        txtCovers.SetFocus
        Exit Sub
    End If
    If errTimer.Enabled = True Then Exit Sub
    If picSlip.Visible = True Then Exit Sub
    Select Case cmdFancy(Index).Caption
        Case "Split Bill"
            txtTables.Text = ""
            txtCovers.Text = ""
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Table to Split"
        Case "Change Waiter"
            '****************** Kotie 22-03-2013
            If UserRecord.Owner_Transfer = False Then
                TillData.UserOveride = 0
                Load frmValidate
                frmValidate.Tag = "Owner Transfer"
                frmValidate.Show vbModal
                DoEvents
                If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                    Validatator = False
                    cmdErr.Caption = "Higher Access Rights required to Change waiter"
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                Else
                    frmValidate.Tag = ""
                    txtTables.Text = ""
                    txtCovers.Text = ""
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    cmdErr.Caption = "Select a Table to Change Owership"
                End If
            End If

        Case "Show All Tables"
            txtTables.Text = ""
            txtCovers.Text = ""
            If UserRecord.All_Tables = True Then
                cmdFancy(Index).Caption = "Show Own Tables"
                LoadTables 1
            Else
                cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "You do not have Access to see all Open Tables"
            End If
         Case "Show Own Tables"
            txtTables.Text = ""
            txtCovers.Text = ""
            If UserRecord.All_Tables = True Then
                cmdFancy(Index).Caption = "Show All Tables"
                LoadTables 0
            Else
                cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "You do not have Access to see all Open Tables"
            End If
        Case "Close Table"
            txtTables.Text = ""
            txtCovers.Text = ""
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Table to Close"
        Case "Transfer Table"
            If UserRecord.Transfers = False Then
             cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "You are not allowed to Transfer Tables"
                Exit Sub
                End If
            txtTables.Text = ""
            txtCovers.Text = ""
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Table to Transfer From"
        Case "Print Bill"
            If TillData.ShortTender = True Then
                cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "You have to finish Tendering to Close the Sale"
                Exit Sub
            End If
            txtTables.Text = ""
            txtCovers.Text = ""
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Table to Print the Bill"
        Case "Split Bill"
            txtTables.Text = ""
            txtCovers.Text = ""
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Table to Split the Bill"
        Case "View Table"
            txtTables.Text = ""
            txtCovers.Text = ""
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Table to View"
        Case "Change Covers"
            txtTables.Text = ""
            txtCovers.Text = ""
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Table to Change the Covers"
        Case "Name Table"
            txtTables.Text = ""
            txtCovers.Text = ""
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Table to Name"
    End Select
    If cmdErr.Visible = True Then Exit Sub
End Sub
Private Sub cmdInput_Click(Index As Integer)
    Select Case cmdInput(Index).Caption
        Case "Reservations"
            If picSlip.Visible = True Then Exit Sub
            If errTimer.Enabled = True Then Exit Sub
            Screen.MousePointer = 11
            Load frmRestRes
            frmRestRes.Tag = "from Tables"
            Screen.MousePointer = 0
            frmRestRes.Show vbModal
            frmRestRes.Tag = ""
        Case "0" To "9"
            If picSlip.Visible = True Then Exit Sub
            If errTimer.Enabled = True And cmdFancy(3).Enabled = True Then Exit Sub
            Select Case cmdInput(13).Tag
                Case "1"
                    If Len(txtTables.Text) <> 4 Then
                        txtTables.Text = txtTables.Text + cmdInput(Index).Caption
                    End If
                Case "2"
                    If Len(txtCovers.Text) <> 4 Then
                        txtCovers.Text = txtCovers.Text + cmdInput(Index).Caption
                    End If
            End Select
        Case "CL"
            If picSlip.Visible = True Then Exit Sub
            If errTimer.Enabled = True And cmdFancy(3).Enabled = True Then
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                cmdErr.Caption = ""
                txtTables.SetFocus
                lblKeyRegister.Tag = ""
                Exit Sub
            End If
            Select Case cmdInput(13).Tag
                Case "1": txtTables.Text = ""
                Case "2": txtCovers.Text = ""
            End Select
        Case "Exit"
            If picSlip.Visible = True Then Exit Sub
            If errTimer.Enabled = True And cmdFancy(3).Enabled = True Then Exit Sub
            If UserRecord.App_Exit = False Then
                Load frmError
                frmError.lblCap.Caption = "You do not have Access Rights to Exit."
                frmError.Show vbModal
                On Error Resume Next
                txtTables.SetFocus
                On Error GoTo 0
                Exit Sub
            Else
                End
            End If
        Case "Ok"
            If txtTables.Enabled = False Then
                If txtCovers.Text = "" Then
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    cmdErr.Caption = "Please provide a value for Covers"
                    cmdInput(13).Tag = "2"
                    Image1.BorderColor = &H80000006
                    txtCovers.Text = ""
                    txtCovers.SetFocus
                    Exit Sub
                End If
                'If Val(txtCovers.Text) < Val(lblKeyRegister.Tag) Then
                  If Val(txtCovers.Text) < 1 Then
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    cmdErr.Caption = "New Covers (" & txtCovers.Text & ") can't be less than the 1" '(" & Val(lblKeyRegister.Tag) & ")"
                    cmdInput(13).Tag = "2"
                    Image1.BorderColor = &H80000006
                    txtCovers.Text = ""
                    txtCovers.SetFocus
                    Exit Sub
                Else
                    ActiveUpdateServer "Update Table_Listing set Covers = " & txtCovers.Text & " where Table_No = " & txtTables.Text
                    Select Case cmdFancy(1).Caption
                        Case "Show All Tables"
                            LoadTables 0
                        Case "Show Own Tables"
                            LoadTables 1
                    End Select
                End If
                txtTables.Enabled = True
                txtTables.Text = ""
                txtCovers.Text = ""
                cmdInput(13).Tag = "1"
                Image2.BorderColor = &H80000006
                cmdErr.Caption = ""
                txtTables.SetFocus
            End If
            If picSlip.Visible = True Then Exit Sub
            If errTimer.Enabled = True And cmdFancy(3).Enabled = True Then Exit Sub
            If Trim(txtTables.Text) <> "" Then
                If Val(txtTables.Text) = 0 Then
                    Load frmError
                    frmError.lblCap.Caption = "You cannot have a Table number Zero. Please change your Table Number."
                    frmError.Show vbModal
                    txtTables.Text = ""
                    txtTables.SetFocus
                    Exit Sub
                End If
                ActiveReadServer "Select * from Table_Listing_View where Table_No = " & Val(txtTables.Text)
                If rs.RecordCount > 0 Then
                    Load frmError
                    If rs.Fields("User_No") = UserRecord.User_Number Then
                        frmError.lblCap.Caption = "Table No " & txtTables.Text & " is already in use by yourself. Please change your Table No."
                    Else
                        frmError.lblCap.Caption = "Table No " & txtTables.Text & " is being used by " & rs.Fields("User_Name") & ". Please change your Table No."
                    End If
                    rs.Close
                    frmError.Show vbModal
                    txtTables.Text = ""
                    txtTables.SetFocus
                    Exit Sub
                End If
                rs.Close
            End If
            Select Case cmdInput(13).Tag
                Case "1"
                    If txtTables.Text = "" Then
                        txtTables.SetFocus
                        Exit Sub
                    End If
                    cmdInput(13).Tag = "2"
                    Image1.BorderColor = &H80000006
                    If cmdFancy(3).Enabled = True Then
                        If txtCovers.Text <> "" Then
                            cmdInput(13).Tag = "1"
                            Image2.BorderColor = &H80000006
                            frmInput.Hide
                            frmSales1.Show
                            frmSales1.cmdDept(6).Caption = "No Sale"
                            TillData.TableNo = txtTables.Text
                            TillData.Covers = txtCovers.Text
                            frmSales1.lblTable = " Table No: " & txtTables.Text & " Covers: " & txtCovers.Text
                            txtTables.Text = ""
                            txtCovers.Text = ""
                            frmSales1.grdMain.Rows = 1
                            frmSales1.lblCash = ""
                            frmSales1.lblTender = "0.00"
                            frmSales1.lblKeyRegister.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
                            Exit Sub
                        Else
                            txtCovers.SetFocus
                        End If
                        txtCovers.SetFocus
                    Else
                        If txtCovers.Text = "" And txtCovers.Enabled = True Then
                            txtCovers.SetFocus
                        Else
                            Screen.MousePointer = 11
                            Load frmServers
                            frmServers.Tag = "frmInput"
                            DoEvents
                            Screen.MousePointer = 0
                            frmServers.Show vbModal
                            Select Case lblServers.Tag
                                Case ""
                                    errTimer.Enabled = False
                                    cmdFancy(1).Enabled = True
                                    cmdFancy(3).Enabled = True
                                    cmdFancy(4).Enabled = True
                                    cmdFancy(5).Enabled = True
                                    cmdFancy(6).Enabled = True
                                    cmdFancy(7).Enabled = True
                                    cmdFancy(8).Enabled = True
                                    txtCovers.Enabled = True
                                    cmdLogoff.Orientation = DIR_NW
                                    Select Case cmdFancy(1).Caption
                                        Case "Show All Tables"
                                            LoadTables 0
                                        Case "Show Own Tables"
                                            LoadTables 1
                                    End Select
                                    txtTables.Enabled = True
                                    txtTables.Text = ""
                                    txtCovers.Text = ""
                                    cmdInput(13).Tag = "1"
                                    Image2.BorderColor = &H80000006
                                    cmdErr.Caption = ""
                                    cmdErr.Visible = False
                                    txtTables.SetFocus
                                Case Else
                                    ActiveReadServer "Select max(User_No) as User_No,Max(Doc_No) as Doc_No from Table_Listing where Table_No = " & lblTransfer.Tag & " group by User_No"
                                    If rs.RecordCount > 0 Then
                                        PreviousOwner = rs.Fields("User_No")
                                        Doc_No = rs.Fields("Doc_No")
                                    Else
                                        Doc_No = 0
                                        PreviousOwner = ""
                                    End If
                                    rs.Close
                                    ActiveUpdateServer "INSERT INTO [Table_Tansfer_Journal]([Date_Time], [Tranfering_User], [Recieving_User], [From_Table], [To_Table], [From_Tab], [To_Tab],[Invoice_No],[User_Action])" & _
                                    " VALUES(Getdate(), " & Val(PreviousOwner) & ", " & lblServers.Tag & ", " & lblTransfer.Tag & ", " & txtTables.Text & ", Null, Null," & Val(Doc_No) & ",'Table to New Table Transfer')"
                                    DoEvents
                                    ActiveUpdateServer "Update Table_Listing set Previous_Owner = User_No, User_No = " & lblServers.Tag & ",Table_no = " & txtTables.Text & " where Table_No = " & lblTransfer.Tag
                                    cmdFancy(1).Enabled = True
                                    cmdFancy(3).Enabled = True
                                    cmdFancy(4).Enabled = True
                                    cmdFancy(5).Enabled = True
                                    cmdFancy(6).Enabled = True
                                    cmdFancy(7).Enabled = True
                                    cmdFancy(8).Enabled = True
                                    txtCovers.Enabled = True
                                    cmdLogoff.Orientation = DIR_NW
                                    Select Case cmdFancy(1).Caption
                                        Case "Show All Tables"
                                            LoadTables 0
                                        Case "Show Own Tables"
                                            LoadTables 1
                                    End Select
                                    cmdErr.Visible = False
                                    cmdErr.BackColor = &HF2&
                                    errTimer.Enabled = False
                                    lblKeyRegister = "Table No: " & lblTransfer.Tag & " transfered to Table No: " & txtTables.Text
                                    lblTransfer.Tag = ""
                                    txtTables.Enabled = True
                                    txtTables.Text = ""
                                    txtCovers.Text = ""
                                    cmdInput(13).Tag = "1"
                                    Image2.BorderColor = &H80000006
                                    cmdErr.Caption = ""
                                    txtTables.SetFocus
                            End Select
                        End If
                    End If
                Case "2"
                    If txtCovers.Text = "" Then
                        cmdInput(13).Tag = "2"
                        txtCovers.SetFocus
                        Exit Sub
                    End If
                    If txtTables.Text = "" Then
                        cmdInput(13).Tag = "1"
                        Image2.BorderColor = &H80000006
                        txtTables.SetFocus
                        Exit Sub
                    End If
                    cmdInput(13).Tag = "1"
                    Image2.BorderColor = &H80000006
                    frmInput.Hide
                    frmSales1.Show
                    TillData.TableNo = txtTables.Text
                    TillData.Covers = txtCovers.Text
                    frmSales1.lblTable = " Table No: " & txtTables.Text & " Covers: " & txtCovers.Text
                    txtTables.Text = ""
                    txtCovers.Text = ""
                    frmSales1.grdMain.Rows = 1
                    frmSales1.lblCash = ""
                    frmSales1.lblTender = "0.00"
                    frmSales1.lblKeyRegister.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
                    Exit Sub
            End Select
    End Select
    Select Case cmdInput(13).Tag
        Case "1": txtTables.SetFocus
        Case "2": txtCovers.SetFocus
    End Select
End Sub

Private Sub cmdLogOff_Click()
    If picSlip.Visible = True Then Exit Sub
    If errTimer.Enabled = True And cmdFancy(3).Enabled = True Then
        Exit Sub
    End If
    errTimer.Enabled = False
    cmdFancy(1).Enabled = True
    cmdFancy(3).Enabled = True
    cmdFancy(4).Enabled = True
    cmdFancy(5).Enabled = True
    cmdFancy(6).Enabled = True
    cmdFancy(7).Enabled = True
    cmdFancy(8).Enabled = True
    txtCovers.Enabled = True
    cmdLogoff.Orientation = DIR_NW
    KeyCode = 0
    'frmSales1.Tag = "Show Splash"
    cmdErr.Caption = ""
    txtTables.Enabled = True
    cmdErr.BackColor = &HF2&
    txtTables.Text = ""
    txtCovers.Text = ""
    cmdErr.Visible = False
    Timer1.Enabled = False
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
    DoEvents
    frmSplash.Show
    frmInput.Hide
    Exit Sub
End Sub
Private Sub cmdTable_Click(Index As Integer)
    
    If picSlip.Visible = True Then Exit Sub
    
    DoEvents
    If cmdTable(Index).Picture = App.Path & "\icons\downArr.bmp" Then
        grdTable.Row = grdTable.Row + 1
        For i = 0 To 19
            If grdTable.TextMatrix(grdTable.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                Else
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                    grdTable.Row = grdTable.Row - 1
                    Exit For
                End If
            Else
                cmdTable(i).Caption = "Table No: " & grdTable.TextMatrix(grdTable.Row, 0)
                cmdTable(i).Tag = grdTable.TextMatrix(grdTable.Row, 1)
                Select Case cmdFancy(1).Caption
                    Case "Show Own Tables"
                        cmdTable(i).TextDescrCB.OffsetY = -10
                        cmdTable(i).TextDescrCB.ColorNormal = &H800000
                        cmdTable(i).TextDescrCB.Text = grdTable.TextMatrix(grdTable.Row, 2)
                    Case "Show All Tables"
                        cmdTable(i).TextDescrCB.Text = ""
                End Select
                If grdTable.TextMatrix(grdTable.Row, 3) = "True" Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    If Trim(grdTable.TextMatrix(grdTable.Row, 5)) = "" Then
                        cmdTable(i).TextDescrCB.Text = ""
                    Else
                        cmdTable(i).TextDescrCB.OffsetY = -10
                        cmdTable(i).TextDescrCB.ColorNormal = &H800000
                        cmdTable(i).TextDescrCB.Text = "From > " & grdTable.TextMatrix(grdTable.Row, 5)
                    End If
                End If
            End If
            If grdTable.Row = grdTable.Rows - 1 Then Exit For
            grdTable.Row = grdTable.Row + 1
        Next i
        For b = i + 1 To cmdTable.Count - 1
            cmdTable(b).Caption = "1"
            cmdTable(b).Tag = ""
            cmdTable(b).Visible = False
        Next b
        Exit Sub
    End If
    If cmdTable(Index).Picture = App.Path & "\icons\upArr.bmp" Then
        cmdTable(0).Picture = ""
        While grdTable.TextMatrix(grdTable.Row, 0) <> "ßArrow"
            grdTable.Row = grdTable.Row - 1
        Wend
        grdTable.Row = grdTable.Row - 19
        For i = 0 To 19
            If grdTable.TextMatrix(grdTable.Row, 0) = "ßArrow" Then
                If i = 0 Then
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\upArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                Else
                    cmdTable(i).Caption = ""
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    cmdTable(i).Picture = App.Path & "\icons\downArr.bmp"
                    If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
                    grdTable.Row = grdTable.Row - 1
                    Exit For
                End If
            Else
                cmdTable(i).Caption = "Table No: " & grdTable.TextMatrix(grdTable.Row, 0)
                cmdTable(i).Tag = grdTable.TextMatrix(grdTable.Row, 1)
                Select Case cmdFancy(1).Caption
                    Case "Show Own Tables"
                        cmdTable(i).TextDescrCB.OffsetY = -10
                        cmdTable(i).TextDescrCB.ColorNormal = &H800000
                        cmdTable(i).TextDescrCB.Text = grdTable.TextMatrix(grdTable.Row, 2)
                    Case "Show All Tables"
                        cmdTable(i).TextDescrCB.Text = ""
                        If Trim(grdTable.TextMatrix(grdTable.Row, 5)) = "" Then
                            cmdTable(i).TextDescrCB.Text = ""
                        Else
                            cmdTable(i).TextDescrCB.OffsetY = -10
                            cmdTable(i).TextDescrCB.ColorNormal = &H800000
                            cmdTable(i).TextDescrCB.Text = "From > " & grdTable.TextMatrix(grdTable.Row, 5)
                        End If
                End Select
                If grdTable.TextMatrix(grdTable.Row, 3) = "True" Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTable(i).TextDescrCT.Text = ""
                End If
                If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
            End If
            If grdTable.Row = grdTable.Rows - 1 Then Exit For
            grdTable.Row = grdTable.Row + 1
        Next i
        For b = i + 1 To cmdTable.Count - 1
            cmdTable(b).Caption = "1"
            cmdTable(b).Tag = ""
            cmdTable(b).TextDescrCB.Text = ""
            cmdTable(b).TextDescrCT.Text = ""
            cmdTable(b).ToolTipText = ""
            cmdTable(b).Visible = False
        Next b
        Exit Sub
    End If
    If cmdTable(Index).Caption = "" Then Exit Sub
    If errTimer.Enabled = True Then
        Select Case cmdErr.Caption
            Case "Select a Table to Print the Bill"
                Panel_no = 1
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                txtTables.SetFocus
                picSlip.Visible = False
                ActiveReadServer "Select * from Table_Listing_view where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                If Workstation_No <> rs.Fields("Workstation_No") Then
                    If rs.Fields("Locked") = True Then
                        cmdErr.Visible = True
                        errTimer.Enabled = True
                        cmdErr.Caption = "This Table is already in opened by another User"
                        cmdTable(Index).TextDescrCT.OffsetY = 12
                        cmdTable(Index).TextDescrCT.ColorNormal = &HC0&
                        cmdTable(Index).TextDescrCT.Text = "In Use"
                        rs.Close
                        Exit Sub
                    End If
                End If
                rs.Close
                TillData.TableNo = Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                frmSales1.LoadOldTable Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                
                '******************* Kotie 20-03-2013
                If TillData.Print_Count > 0 Then
                    If UserRecord.Reprint = False Then
                        TillData.UserOveride = 0
                        Load frmValidate
                        frmValidate.Tag = "Reprint"
                        frmValidate.Show vbModal
                        If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                            ActiveUpdateServer "Update Table_Listing set Locked =0 where Table_No= " & TillData.TableNo
                            frmInput.errTimer = True
                            frmInput.cmdErr.Caption = "Not allowed"
                            cmdErr.Visible = True
                            Exit Sub
                        Else
                      
                        End If
                    End If
                End If
                '********************
                
                DoEvents
                lblKeyRegister = "Printed Bill for Table No: " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) & " Covers: " & cmdTable(Index).Tag
                LoadTable Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                ActiveUpdateServer ("Insert into Print_Journal (User_No,doc_No,Doc_Type,DateTimePrinted, User_override, Table_no)VALUES(" & UserRecord.User_Number & "," & TillData.DocNo & ",'Bill Print', getdate(), '" & TillData.UserOveride & "','" & TillData.TableNo & "')")
                PrintSlip "Print Bill Table"
                frmInput.cmdErr.Caption = ""
                With frmSales1
                    ActiveUpdateServer "Update Table_Listing set Locked =0 where Table_No= " & TillData.TableNo
                    .lblTable = ""
                    .grdMain.Rows = 1
                    .lblCash.Caption = ""
                    .lblTender.Caption = "0.00"
                    TillData.DocNo = 0
                    TillData.TableNo = 0
                    TillData.Table_Name = ""
                    TillData.Covers = 0
                    TillData.TotDiscount = 0
                    TillData.TotDiscountVal = 0
                    TillData.TotDiscountCount = 0
                    TillData.TotDiscountValCount = 0
                    GlobalMode = TillMode.FinMode
                    DoEvents
                    Exit Sub
                End With
            Case "Select a Table to Change Owership"
                If cmdTable(Index).TextDescrCT.Text = "In Use" Then
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    Select Case cmdFancy(1).Caption
                        Case "Show Own Tables"
                            If UserRecord.User_Number <> Val(Trim(Mid(cmdTable(Index).TextDescrCB.Text, 1, InStr(cmdTable(Index).TextDescrCB.Text, "-") - 1))) Then
                                cmdErr.Caption = "This Table is already in opened by another User"
                            Else
                                cmdErr.Caption = "This Table is already open on another Workstation"
                            End If
                        Case "Show All Tables"
                            ActiveReadServer "Select * from Table_Listing where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                            If rs.RecordCount > 0 Then
                                If rs.Fields("User_No") = UserRecord.User_Number Then
                                    If Workstation_No = rs.Fields("Workstation_No") Then
                                        GoTo skip1
                                    End If
                                End If
                            End If
                            rs.Close
                            cmdErr.Caption = "This Table is already open on another Workstation"
                    End Select
                    Exit Sub
                End If
                Screen.MousePointer = 11
                Load frmServers
                frmServers.Tag = "frmInput"
                DoEvents
                Screen.MousePointer = 0
                frmServers.Show vbModal
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                If lblServers.Tag <> "" Then
                    ActiveReadServer "Select max(User_No) as User_No, Max(Doc_No) as Doc_No from Table_Listing where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) & " group by User_No"
                    If rs.RecordCount > 0 Then
                        PreviousOwner = rs.Fields("User_No")
                        Doc_No = rs.Fields("Doc_No")
                    Else
                        PreviousOwner = ""
                        Doc_No = 0
                    End If
                    rs.Close
                    ActiveUpdateServer "INSERT INTO [Table_Tansfer_Journal]([Date_Time], [Tranfering_User], [Recieving_User], [From_Table], [To_Table], [From_Tab], [To_Tab],[Invoice_No],[User_Action])" & _
                    " VALUES(Getdate(), " & Val(PreviousOwner) & ", " & lblServers.Tag & "," & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) & ", " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) & ", Null, Null," & Val(Doc_No) & ",'Waiter Change')"
                    ActiveUpdateServer "Update Table_Listing set Previous_Owner = User_No, User_No = " & lblServers.Tag & " where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                End If
                Select Case cmdFancy(1).Caption
                    Case "Show All Tables"
                        LoadTables 0
                    Case "Show Own Tables"
                        LoadTables 1
                End Select
            Case "Select a Table to Transfer to or Start a New Table"
                If cmdTable(Index).TextDescrCT.Text = "In Use" Then
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    Select Case cmdFancy(1).Caption
                        Case "Show Own Tables"
                            If UserRecord.User_Number <> Val(Trim(Mid(cmdTable(Index).TextDescrCB.Text, 1, InStr(cmdTable(Index).TextDescrCB.Text, "-") - 1))) Then
                                cmdErr.Caption = "This Table is already in opened by another User"
                            Else
                                cmdErr.Caption = "This Table is already open on another Workstation"
                            End If
                        Case "Show All Tables"
                            ActiveReadServer "Select * from Table_Listing where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                            If rs.RecordCount > 0 Then
                                If rs.Fields("User_No") = UserRecord.User_Number Then
                                    If Workstation_No = rs.Fields("Workstation_No") Then
                                        GoTo skip1
                                    End If
                                End If
                            End If
                            rs.Close
                            cmdErr.Caption = "This Table is already open on another Workstation"
                    End Select
                    Exit Sub
                End If
                ActiveReadServer "Select max(User_No) as User_No, Max(Doc_No) as Doc_No from Table_Listing where Table_No = " & lblTransfer.Tag & " group by User_No"
                If rs.RecordCount > 0 Then
                    PreviousOwner = rs.Fields("User_No")
                    Doc_No = rs.Fields("Doc_No")
                Else
                    PreviousOwner = ""
                    Doc_No = 0
                End If
                rs.Close
                
                
                
                For i = 0 To grdTable.Rows - 1
                    If Val(grdTable.TextMatrix(i, 0)) <> "9999" Then
                    If Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) = Val(grdTable.TextMatrix(i, 0)) Then
                        NewUser = Val(Mid(grdTable.TextMatrix(i, 2), 1, InStr(grdTable.TextMatrix(i, 2), "-") - 1))
                        Exit For
                    End If
                End If
                Next i
                ActiveUpdateServer "INSERT INTO [Table_Tansfer_Journal]([Date_Time], [Tranfering_User], [Recieving_User], [From_Table], [To_Table], [From_Tab], [To_Tab],[Invoice_No],[User_Action])" & _
                " VALUES(Getdate(), " & Val(PreviousOwner) & ", " & NewUser & "," & lblTransfer.Tag & ", " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) & ", Null, Null," & Val(Doc_No) & ",'Table to Table Transfer')"
                DoEvents
                
                ActiveUpdateServer "Update Table_Listing set User_No = " & NewUser & ",Covers = Covers +  " & cmdTable(Index).Tag & ",Table_no = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) & " where Table_No = " & lblTransfer.Tag
                ActiveUpdateServer "Update Table_Listing set Previous_Owner = '" & PreviousOwner & "' where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                ActiveUpdateServer "Update Table_Listing set Table_name = '" & cmdTable(Index).TextDescrLT.Text & "'  where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                
                
                cmdFancy(1).Enabled = True
                cmdFancy(3).Enabled = True
                cmdFancy(4).Enabled = True
                cmdFancy(5).Enabled = True
                cmdFancy(6).Enabled = True
                cmdFancy(7).Enabled = True
                cmdFancy(8).Enabled = True
                txtCovers.Enabled = True
                cmdLogoff.Orientation = DIR_NW
                Select Case cmdFancy(1).Caption
                    Case "Show All Tables"
                        LoadTables 0
                    Case "Show Own Tables"
                        LoadTables 1
                End Select
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                lblKeyRegister = "Table No: " & lblTransfer.Tag & " transfered to Table No: " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                lblTransfer.Tag = ""
                Exit Sub
            Case "Select a Table to Transfer From"
                If cmdTable(Index).TextDescrCT.Text = "In Use" Then
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    Select Case cmdFancy(1).Caption
                        Case "Show Own Tables"
                            If UserRecord.User_Number <> Val(Trim(Mid(cmdTable(Index).TextDescrCB.Text, 1, InStr(cmdTable(Index).TextDescrCB.Text, "-") - 1))) Then
                                cmdErr.Caption = "This Table is already in opened by another User"
                            Else
                                cmdErr.Caption = "This Table is already open on another Workstation"
                            End If
                        Case "Show All Tables"
                            ActiveReadServer "Select * from Table_Listing where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                            If rs.RecordCount > 0 Then
                                If rs.Fields("User_No") = UserRecord.User_Number Then
                                    If Workstation_No = rs.Fields("Workstation_No") Then
                                        GoTo skip1
                                    End If
                                End If
                            End If
                            rs.Close
                            cmdErr.Caption = "This Table is already open on another Workstation"
                    End Select
                    Exit Sub
                End If
                cmdFancy(1).Enabled = False
                cmdFancy(3).Enabled = False
                cmdFancy(4).Enabled = False
                cmdFancy(5).Enabled = False
                cmdFancy(6).Enabled = False
                cmdFancy(7).Enabled = False
                cmdFancy(8).Enabled = False
                txtCovers.Enabled = False
                cmdLogoff.Orientation = DIR_WEST
                txtTables.SetFocus
                cmdErr.Caption = "Select a Table to Transfer to or Start a New Table"
                lblTransfer.Caption = "Transfering Table No: " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) & " with " & cmdTable(Index).Tag & " Covers"
                lblTransfer.Tag = Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                Select Case cmdFancy(1).Caption
                    Case "Show All Tables"
                        LoadTables 2
                    Case "Show Own Tables"
                        LoadTables 3
                End Select
                Exit Sub
            Case "Select a Table to Change the Covers"
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                txtTables.Text = Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                For i = 0 To grdTable.Rows - 1
                    If Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) = Val(grdTable.TextMatrix(i, 0)) Then
                        lblKeyRegister.Tag = Val(grdTable.TextMatrix(i, 1))
                        Select Case grdTable.TextMatrix(i, 1)
                            Case 1
                                lblKeyRegister.Caption = "Table No: " & Val(grdTable.TextMatrix(i, 0)) & " currently has " & grdTable.TextMatrix(i, 1) & " Cover"
                            Case Else
                                lblKeyRegister.Caption = "Table No: " & Val(grdTable.TextMatrix(i, 0)) & " currently has " & grdTable.TextMatrix(i, 1) & " Covers"
                        End Select
                        Exit For
                    End If
                Next i
                txtTables.Enabled = False
                cmdInput(13).Tag = "2"
                Image1.BorderColor = &H80000006
                txtCovers.SetFocus
            Case "Select a Table to Close"
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                ActiveReadServer "Select * from Table_Listing_view where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                If Workstation_No <> rs.Fields("Workstation_No") Then
                    If rs.Fields("Locked") = True Then
                        cmdErr.Visible = True
                        errTimer.Enabled = True
                        cmdErr.Caption = "This Table is already in opened by another User"
                        cmdTable(Index).TextDescrCT.OffsetY = 12
                        cmdTable(Index).TextDescrCT.ColorNormal = &HC0&
                        cmdTable(Index).TextDescrCT.Text = "In Use"
                        rs.Close
                        Exit Sub
                    End If
                End If
                rs.Close
                TillData.TableNo = Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                frmSales1.LoadOldTable Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                Panel_no = 1
                Key_Function ("Close Table")
            Case "Select a Table to Split"
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                If Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) <> Int(Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))) Then
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    cmdErr.Caption = "Invalid Key Pressed"
                    Exit Sub
                End If
                ActiveReadServer "Select * from Table_Listing_view where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                If Workstation_No <> rs.Fields("Workstation_No") Then
                    If rs.Fields("Locked") = True Then
                        cmdErr.Visible = True
                        errTimer.Enabled = True
                        cmdErr.Caption = "This Table is already in opened by another User"
                        cmdTable(Index).TextDescrCT.OffsetY = 12
                        cmdTable(Index).TextDescrCT.ColorNormal = &HC0&
                        cmdTable(Index).TextDescrCT.Text = "In Use"
                        rs.Close
                        Exit Sub
                    End If
                End If
                rs.Close
                TillData.TableNo = Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                frmSales1.LoadOldTable Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                Panel_no = 1
                Key_Function ("Split Bill")
            Case "Select a Table to View"
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                txtTables.SetFocus
                picSlip.Visible = True
                lblKeyRegister = "Viewing Table No: " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)) & " Covers: " & cmdTable(Index).Tag
                LoadTable Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                grdMain.SetFocus
                Exit Sub
                'Kotie 22-03-2013  14:35
            Case "Select a Table to Name"
                frmKeyBoard.Tag = "Tables"
                frmKeyBoard.Show vbModal
                TillData.Table_Name = frmKeyBoard.txtReg
                cmdTable(Index).TextDescrLT.Text = TillData.Table_Name
                ActiveUpdateServer ("Update Table_Listing set Table_name = '" & TillData.Table_Name & "' where Table_No= " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1)))
                TillData.Table_Name = ""
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
        End Select
    Else
        If picSlip.Visible = True Then Exit Sub
        
        If cmdTable(Index).TextDescrCT.Text = "In Use" Then
            cmdErr.Visible = True
            errTimer.Enabled = True
            Select Case cmdFancy(1).Caption
                Case "Show Own Tables"
                    If UserRecord.User_Number <> Val(Trim(Mid(cmdTable(Index).TextDescrCB.Text, 1, InStr(cmdTable(Index).TextDescrCB.Text, "-") - 1))) Then
                        cmdErr.Caption = "This Table is already in opened by another User"
                    Else
                        cmdErr.Caption = "This Table is already open on another Workstation"
                    End If
                Case "Show All Tables"
                    ActiveReadServer "Select * from Table_Listing where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
                    If rs.RecordCount > 0 Then
                        If rs.Fields("User_No") = UserRecord.User_Number Then
                            If Workstation_No = rs.Fields("Workstation_No") Then
                                GoTo skip1
                            End If
                        End If
                    End If
                    rs.Close
                    cmdErr.Caption = "This Table is already open on another Workstation"
            End Select
            Exit Sub
        End If
skip1:
        If txtTables.Enabled = False Then
            txtCovers.SetFocus
            Exit Sub
        End If
        ActiveReadServer "Select * from Table_Listing_view where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
        
        If Workstation_No <> rs.Fields("Workstation_No") Then
            If rs.Fields("Locked") = True Then 'And rs.Fields("User_no") <> UserRecord.User_Number Then
                cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "This Table is already in Open on Workstation " & rs.Fields("Workstation_No") & " ."
                cmdTable(Index).TextDescrCT.OffsetY = 12
                cmdTable(Index).TextDescrCT.ColorNormal = &HC0&
                cmdTable(Index).TextDescrCT.Text = "In Use"
                rs.Close
                Exit Sub
            End If
        End If
        rs.Close
        ActiveUpdateServer " Update Table_Listing set Locked = '1',Workstation_No = '" & Workstation_No & "'  where Table_No = " & Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
        TillData.TableNo = Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
        frmSales1.LoadOldTable Val(Mid(cmdTable(Index).Caption, InStr(cmdTable(Index).Caption, ":") + 1))
    End If
End Sub
Private Sub LoadTable(Table_Number)
    ActiveReadServer "Select * from Table_Listing where Table_No= " & Table_Number & " order by Line_No"
    grdMain.Rows = 1
    grdMain.ColHidden(14) = True
    lblWorkstation.Caption = "Opened on Workstation No: " & rs.Fields("Workstation_No")
    If rs.RecordCount > 0 Then
        ActiveReadServer2 "Select Date_Time from Sales_Journal where Invoice_No = " & rs.Fields("Doc_No")
        If rs2.RecordCount > 0 Then
            lblDateOpened.Caption = "Table Opened on " & Format(rs2.Fields("Date_Time"), "DDD DD MMM YYYY HH:MM")
        End If
        rs2.Close
    End If
    While Not rs.EOF
        grdMain.Rows = grdMain.Rows + 1
        grdMain.TextMatrix(grdMain.Rows - 1, 0) = Trim(rs.Fields("Qty") & "")
        grdMain.TextMatrix(grdMain.Rows - 1, 1) = rs.Fields("Short_Desc")
        If Val(rs.Fields("Line_Total") & "") <> 0 Then
            grdMain.TextMatrix(grdMain.Rows - 1, 2) = Format(rs.Fields("Line_Total"), "0.00")
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 3) = rs.Fields("KeyString")
        If Trim(rs.Fields("KeyString") & "") = "Subtotal" Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0FFC0
        End If
        If Trim(rs.Fields("KeyString") & "") = "" Then
            grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC00000
            grdMain.Cell(flexcpFontBold, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = True
        End If
        If grdMain.ValueMatrix(grdMain.Rows - 1, 2) <> 0 Then
            grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &H0
            grdMain.Cell(flexcpFontBold, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = True
        End If
        
        grdMain.TextMatrix(grdMain.Rows - 1, 4) = rs.Fields("Cost")
        grdMain.TextMatrix(grdMain.Rows - 1, 5) = rs.Fields("Tax_Rate")
        grdMain.TextMatrix(grdMain.Rows - 1, 6) = rs.Fields("Tax_Type")
        grdMain.TextMatrix(grdMain.Rows - 1, 7) = Trim(rs.Fields("Keyregister") & "")
        grdMain.TextMatrix(grdMain.Rows - 1, 8) = Trim(rs.Fields("Extra_Function") & "")
        If rs.Fields("Extra_Function") & "" <> "" Then
            Select Case rs.Fields("Extra_Function") & ""
                Case "Void"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0C0FF
                Case "Corr"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0C0FF
                    grdMain.Cell(flexcpFontStrikethru, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = True
                Case "Return Item"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0C0FF
                Case "Wastage"
                    grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = &HC0C0FF
                Case Else
                    If rs.Fields("Extra_Function") & "" <> "" Then
                        grdMain.Cell(flexcpForeColor, grdMain.Rows - 1, 0, grdMain.Rows - 1, 2) = vbRed
                    End If
            End Select
        End If
        grdMain.TextMatrix(grdMain.Rows - 1, 9) = rs.Fields("Product_Code")
        grdMain.TextMatrix(grdMain.Rows - 1, 10) = rs.Fields("Dept_No")
        grdMain.TextMatrix(grdMain.Rows - 1, 11) = rs.Fields("Kitchen1") & ""
        grdMain.TextMatrix(grdMain.Rows - 1, 12) = rs.Fields("Kitchen2") & "'"
        grdMain.TextMatrix(grdMain.Rows - 1, 13) = rs.Fields("Price_Override")
        grdMain.TextMatrix(grdMain.Rows - 1, 14) = rs.Fields("Printed") & ""
        If rs.Fields("Printed") & "" = "P" Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 14, grdMain.Rows - 1, 14) = &HC0FFFF
            grdMain.ColHidden(14) = False
        End If
        If rs.Fields("Printed") & "" = Chr(187) Then
            grdMain.Cell(flexcpBackColor, grdMain.Rows - 1, 14, grdMain.Rows - 1, 14) = &HC0FFC0
            grdMain.ColHidden(14) = False
        End If
        rs.MoveNext
    Wend
    rs.Close
    grdMain.Row = grdMain.Rows - 1
    grdMain.ShowCell grdMain.Rows - 1, 0
    Sale_Total = 0
    For i = 1 To grdMain.Rows - 1
        If grdMain.TextMatrix(i, 8) <> "Corr" Then
            If grdMain.TextMatrix(i, 3) <> "Subtotal" Then
                Sale_Total = Val(Sale_Total) + Val(grdMain.TextMatrix(i, 2))
            End If
        End If
    Next i
    lblTender.Caption = Format(Sale_Total, "0.00")
    ActiveReadServer "Update Table_Listing set Previous_Owner = null where Table_No= " & Table_Number
End Sub
Private Sub errTimer_Timer()
    Select Case cmdErr.BackColor
        Case &HF2&      'White
            cmdErr.BackColor = &HFFFF&
        Case &HFFFF&    'Yellow
            cmdErr.BackColor = &HF2&
    End Select
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = 0
    If Me.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.782
            Me.Controls(i).Left = Me.Controls(i).Left * 0.778
            Me.Controls(i).Height = Me.Controls(i).Height * 0.782
            Me.Controls(i).top = Me.Controls(i).top * 0.782
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
        On Error GoTo 0
        newBack.Width = Me.Width
        newBack.Height = Me.Height
    End If
    If TillData.TableNo = 0 And cmdFancy(3).Enabled = True Then
        lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
        lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
        cmdErr.Caption = ""
        txtTables.SetFocus
        cmdFancy(1).Caption = "Show All Tables"
        LoadTables 0
    End If
    lblKeyRegister.Tag = ""
    On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static DallasString
    Select Case KeyCode
        Case 48
        Case 13
            If Left(DallasString, 2) = "ßß" Then
                DallasString = ""
                cmdLogOff_Click
            Else
                cmdInput_Click 13
                DallasString = ""
            End If
        Case 100
            DallasString = DallasString & "ß"
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48
            cmdInput_Click 10
        Case 49
            cmdInput_Click 0
        Case 50
            cmdInput_Click 1
        Case 51
            cmdInput_Click 2
        Case 52
            cmdInput_Click 3
        Case 53
            cmdInput_Click 4
        Case 54
            cmdInput_Click 5
        Case 55
            cmdInput_Click 6
        Case 56
            cmdInput_Click 7
        Case 57
            cmdInput_Click 8
    End Select
    KeyAscii = 0
End Sub

Private Sub Label3_Click()

End Sub

Private Sub scrolTimer_Timer()
    scrolTimer.Interval = 50
    Select Case scrolTimer.Tag
        Case "0"
            If grdMain.Row <> 1 Then
                grdMain.Row = grdMain.Row - 1
            End If
        Case "1"
            If grdMain.Row <> grdMain.Rows - 1 Then
                grdMain.Row = grdMain.Row + 1
            End If
    End Select
    grdMain.ShowCell grdMain.Row, 0
End Sub
Private Sub cmdArrow_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    scrolTimer.Tag = Index
    scrolTimer.Enabled = True
End Sub
Private Sub cmdArrow_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    scrolTimer.Enabled = False
    scrolTimer.Interval = 1000
End Sub
Private Sub Form_Load() '
    picSlip.Width = 6495
    grdMain.TextMatrix(0, 0) = " No"
    grdMain.TextMatrix(0, 1) = "Description"
    grdMain.TextMatrix(0, 2) = "Total "
    grdMain.ColWidth(0) = 500
    grdMain.ColWidth(1) = 3250
    grdMain.ColWidth(14) = 100
    grdMain.ColAlignment(0) = flexAlignLeftCenter
    grdMain.ColAlignment(1) = flexAlignLeftCenter
    grdMain.ColAlignment(2) = flexAlignRightCenter
    grdMain.ColHidden(3) = True
    grdMain.ColHidden(4) = True
    grdMain.ColHidden(5) = True
    grdMain.ColHidden(6) = True
    grdMain.ColHidden(7) = True
    grdMain.ColHidden(8) = True
    grdMain.ColHidden(9) = True
    grdMain.ColHidden(10) = True
    grdMain.ColHidden(11) = True
    grdMain.ColHidden(12) = True
    grdMain.ColHidden(13) = True
    grdMain.ColHidden(14) = True
    grdMain.ColHidden(15) = True
    grdMain.ColHidden(16) = True
    grdMain.Cell(flexcpForeColor, 0, 0, 0, 2) = &H808080
End Sub
Private Sub Image4_Click()
    If txtCovers.Enabled = True Then
        cmdInput(13).Tag = "2"
        Image1.BorderColor = &H80000006
        txtCovers.SetFocus
    End If
End Sub
Private Sub Image3_Click()
    If txtTables.Enabled = True Then
        cmdInput(13).Tag = "1"
        Image2.BorderColor = &H80000006
        txtTables.SetFocus
    End If
End Sub

Private Sub Timer1_Timer()
    Select Case cmdInput(13).Tag
        Case "1"
            Select Case Image1.BorderColor
                Case &H80000006
                    Image1.BorderColor = &HFF&
                Case &HFF&
                    Image1.BorderColor = &H80000006
            End Select
        Case "2"
            Select Case Image2.BorderColor
                Case &H80000006
                    Image2.BorderColor = &HFF&
                Case &HFF&
                    Image2.BorderColor = &H80000006
            End Select
    End Select
End Sub
Private Sub LoadTables(Action)
    grdTable.Rows = 0
    grdTable.Cols = 7
    cmdTable(0).Caption = ""
    cmdTable(0).Picture = ""
    DoEvents
    Select Case Action
        Case 0
            ActiveReadServer "Select * from Table_Listing_View where User_No = " & UserRecord.User_Number
        Case 1
            ActiveReadServer "Select * from Table_Listing_View"
        Case 2
            ActiveReadServer "Select * from Table_Listing_View where User_No = " & UserRecord.User_Number & " and Table_No <> " & lblTransfer.Tag
        Case 3
            ActiveReadServer "Select * from Table_Listing_View where Table_No <> " & lblTransfer.Tag
    End Select
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdTable.Rows = grdTable.Rows + 1
        If i < 19 And Not rs.EOF Then
            
            cmdTable(i).TextDescrLT.Text = rs.Fields("Table_Name") & ""
            cmdTable(i).Caption = "Table No: " & rs.Fields("Table_No")
            If rs.Fields("Table_No") = 9999 Then
            cmdTable(i).FontTextCaption.Size = 12
            cmdTable(i).ForeColor = vbRed
            cmdTable(i).Caption = rs.Fields("Table_No") & " Training Table"
            End If
            If rs.Fields("Table_No") <> 9999 Then
            cmdTable(i).ForeColor = vbBlack
            cmdTable(i).FontTextCaption.Size = 16
            End If
            cmdTable(i).Tag = rs.Fields("Covers")
            If cmdTable(i).Visible = False Then cmdTable(i).Visible = True
            grdTable.Row = grdTable.Rows - 1
            grdTable.TextMatrix(grdTable.Rows - 1, 0) = rs.Fields("Table_No")
            grdTable.TextMatrix(grdTable.Rows - 1, 1) = rs.Fields("Covers")
            grdTable.TextMatrix(grdTable.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTable.TextMatrix(grdTable.Rows - 1, 3) = rs.Fields("Locked")
            grdTable.TextMatrix(grdTable.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTable.TextMatrix(grdTable.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            grdTable.TextMatrix(grdTable.Rows - 1, 6) = rs.Fields("Table_name") & ""
            If Action = 1 Or Action = 3 Then
                cmdTable(i).TextDescrCB.OffsetY = -10
                cmdTable(i).TextDescrCB.ColorNormal = &H800000
                cmdTable(i).TextDescrCB.Text = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                If rs.Fields("Locked") = True Then
                    cmdTable(i).TextDescrCT.OffsetY = 12
                    cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTable(i).TextDescrCT.Text = ""
                End If
            Else
                cmdTable(i).TextDescrCB.Text = ""
                If rs.Fields("Previous_Owner") = rs.Fields("User_No") Then
                        If rs.Fields("Locked") = True Then
                            cmdTable(i).TextDescrCT.OffsetY = 12
                            cmdTable(i).TextDescrCT.ColorNormal = &HC0&
                            cmdTable(i).TextDescrCT.Text = "In Use"
                        Else
                            cmdTable(i).TextDescrCT.Text = ""
                        End If
                Else
                    cmdTable(i).TextDescrCB.OffsetY = -10
                    cmdTable(i).TextDescrCB.ColorNormal = &H800000
                    If rs.Fields("Previous_Name") & "" <> "" Then
                        If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                            cmdTable(i).TextDescrCB.Text = "From > " & rs.Fields("Previous_Name") & ""
                        Else
                            cmdTable(i).TextDescrCB.Text = ""
                        End If
                    End If
                End If
                If rs.Fields("Locked") = True Then
                    cmdTable(i).TextDescrCB.OffsetY = 12
                    cmdTable(i).TextDescrCB.ColorNormal = &HC0&
                    cmdTable(i).TextDescrCB.Text = "In Use"
                Else
                  '  cmdTable(i).TextDescrCB.Text = ""
                End If
            End If
        Else
            If b = 0 Then
                grdTable.TextMatrix(grdTable.Rows - 1, 0) = "ßArrow"
                grdTable.Rows = grdTable.Rows + 1
                If i = 19 Then
                    cmdTable(19).Caption = ""
                    cmdTable(19).Picture = App.Path & "\icons\downArr.bmp"
                    cmdTable(i).TextDescrCB.Text = ""
                    cmdTable(i).TextDescrCT.Text = ""
                    If cmdTable(19).Visible = False Then cmdTable(19).Visible = True
                End If
            End If
            b = b + 1
            
            grdTable.TextMatrix(grdTable.Rows - 1, 0) = rs.Fields("Table_No")
            
            
            
            grdTable.TextMatrix(grdTable.Rows - 1, 1) = rs.Fields("Covers")
            grdTable.TextMatrix(grdTable.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTable.TextMatrix(grdTable.Rows - 1, 3) = rs.Fields("Locked")
            grdTable.TextMatrix(grdTable.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTable.TextMatrix(grdTable.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            If b = 18 Then b = 0
        End If
        rs.MoveNext
    Wend
    Select Case Action
        Case 0, 2
            If rs.RecordCount = 1 Then
                lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Table"
            Else
                lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Tables"
            End If
        Case 1, 3
            If rs.RecordCount = 1 Then
                lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Table"
            Else
                lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Tables"
            End If
    End Select
    If rs.RecordCount > 0 Then
        lblTables.Visible = False
    Else
        lblTables.Visible = True
    End If
    rs.Close
    For b = i + 1 To cmdTable.Count - 1
       cmdTable(b).Caption = "0"
       cmdTable(b).Visible = False
    Next b
End Sub

Private Sub Timer2_Timer()
    For i = 0 To 2
        If shpLive(i).BackColor = &HFF00& Then
            shpLive(i).BackColor = &HFFFFFF
            If i = 2 Then
                    shpLive(0).BackColor = &HFF00&
                Else
                    shpLive(i + 1).BackColor = &HFF00&
            End If
            Exit For
        End If
    Next i
End Sub

Private Sub Timer3_Timer()
     lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
End Sub

