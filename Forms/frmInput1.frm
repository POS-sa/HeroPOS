VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmInput1 
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInput1.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin BTNENHLib4.BtnEnh lblTransfer 
      Height          =   1265
      Left            =   360
      TabIndex        =   66
      Top             =   2340
      Visible         =   0   'False
      Width           =   9390
      _Version        =   524298
      _ExtentX        =   16563
      _ExtentY        =   2231
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
      dibPicture      =   "frmInput1.frx":E4EF
      textCaption     =   "frmInput1.frx":F415
      textLT          =   "frmInput1.frx":F48F
      textCT          =   "frmInput1.frx":F4A7
      textRT          =   "frmInput1.frx":F4BF
      textLM          =   "frmInput1.frx":F4D7
      textRM          =   "frmInput1.frx":F4EF
      textLB          =   "frmInput1.frx":F507
      textCB          =   "frmInput1.frx":F51F
      textRB          =   "frmInput1.frx":F537
      colorBack       =   "frmInput1.frx":F54F
      colorIntern     =   "frmInput1.frx":F579
      colorMO         =   "frmInput1.frx":F5A3
      colorFocus      =   "frmInput1.frx":F5CD
      colorDisabled   =   "frmInput1.frx":F5F7
      colorPressed    =   "frmInput1.frx":F621
      Style           =   7
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
      SurfaceTransparentZone=   1
   End
   Begin VB.PictureBox picTables 
      BackColor       =   &H00B6DBED&
      ForeColor       =   &H00B6DBED&
      Height          =   9195
      Left            =   9870
      ScaleHeight     =   9135
      ScaleWidth      =   5145
      TabIndex        =   47
      Top             =   1380
      Visible         =   0   'False
      Width           =   5205
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1425
         Index           =   11
         Left            =   3420
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   7635
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
         _ExtentY        =   2514
         _StockProps     =   66
         Caption         =   "Cancel"
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
         textCaption     =   "frmInput1.frx":F64B
         textLT          =   "frmInput1.frx":F6B7
         textCT          =   "frmInput1.frx":F6CF
         textRT          =   "frmInput1.frx":F6E7
         textLM          =   "frmInput1.frx":F6FF
         textRM          =   "frmInput1.frx":F717
         textLB          =   "frmInput1.frx":F72F
         textCB          =   "frmInput1.frx":F747
         textRB          =   "frmInput1.frx":F75F
         colorBack       =   "frmInput1.frx":F777
         colorIntern     =   "frmInput1.frx":F7A1
         colorMO         =   "frmInput1.frx":F7CB
         colorFocus      =   "frmInput1.frx":F7F5
         colorDisabled   =   "frmInput1.frx":F81F
         colorPressed    =   "frmInput1.frx":F849
         Orientation     =   7
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   8
         Left            =   3420
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   6270
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":F873
         textLT          =   "frmInput1.frx":F8D5
         textCT          =   "frmInput1.frx":F8ED
         textRT          =   "frmInput1.frx":F905
         textLM          =   "frmInput1.frx":F91D
         textRM          =   "frmInput1.frx":F935
         textLB          =   "frmInput1.frx":F94D
         textCB          =   "frmInput1.frx":F965
         textRB          =   "frmInput1.frx":F97D
         colorBack       =   "frmInput1.frx":F995
         colorIntern     =   "frmInput1.frx":F9BF
         colorMO         =   "frmInput1.frx":F9E9
         colorFocus      =   "frmInput1.frx":FA13
         colorDisabled   =   "frmInput1.frx":FA3D
         colorPressed    =   "frmInput1.frx":FA67
         Orientation     =   6
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   2
         Left            =   3420
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
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
         textCaption     =   "frmInput1.frx":FA91
         textLT          =   "frmInput1.frx":FAF3
         textCT          =   "frmInput1.frx":FB0B
         textRT          =   "frmInput1.frx":FB23
         textLM          =   "frmInput1.frx":FB3B
         textRM          =   "frmInput1.frx":FB53
         textLB          =   "frmInput1.frx":FB6B
         textCB          =   "frmInput1.frx":FB83
         textRB          =   "frmInput1.frx":FB9B
         colorBack       =   "frmInput1.frx":FBB3
         colorIntern     =   "frmInput1.frx":FBDD
         colorMO         =   "frmInput1.frx":FC07
         colorFocus      =   "frmInput1.frx":FC31
         colorDisabled   =   "frmInput1.frx":FC5B
         colorPressed    =   "frmInput1.frx":FC85
         Orientation     =   6
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   5
         Left            =   3420
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4905
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":FCAF
         textLT          =   "frmInput1.frx":FD11
         textCT          =   "frmInput1.frx":FD29
         textRT          =   "frmInput1.frx":FD41
         textLM          =   "frmInput1.frx":FD59
         textRM          =   "frmInput1.frx":FD71
         textLB          =   "frmInput1.frx":FD89
         textCB          =   "frmInput1.frx":FDA1
         textRB          =   "frmInput1.frx":FDB9
         colorBack       =   "frmInput1.frx":FDD1
         colorIntern     =   "frmInput1.frx":FDFB
         colorMO         =   "frmInput1.frx":FE25
         colorFocus      =   "frmInput1.frx":FE4F
         colorDisabled   =   "frmInput1.frx":FE79
         colorPressed    =   "frmInput1.frx":FEA3
         Orientation     =   6
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1425
         Index           =   10
         Left            =   1785
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   7635
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":FECD
         textLT          =   "frmInput1.frx":FF2F
         textCT          =   "frmInput1.frx":FF47
         textRT          =   "frmInput1.frx":FF5F
         textLM          =   "frmInput1.frx":FF77
         textRM          =   "frmInput1.frx":FF8F
         textLB          =   "frmInput1.frx":FFA7
         textCB          =   "frmInput1.frx":FFBF
         textRB          =   "frmInput1.frx":FFD7
         colorBack       =   "frmInput1.frx":FFEF
         colorIntern     =   "frmInput1.frx":10019
         colorMO         =   "frmInput1.frx":10043
         colorFocus      =   "frmInput1.frx":1006D
         colorDisabled   =   "frmInput1.frx":10097
         colorPressed    =   "frmInput1.frx":100C1
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   7
         Left            =   1785
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   6270
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":100EB
         textLT          =   "frmInput1.frx":1014D
         textCT          =   "frmInput1.frx":10165
         textRT          =   "frmInput1.frx":1017D
         textLM          =   "frmInput1.frx":10195
         textRM          =   "frmInput1.frx":101AD
         textLB          =   "frmInput1.frx":101C5
         textCB          =   "frmInput1.frx":101DD
         textRB          =   "frmInput1.frx":101F5
         colorBack       =   "frmInput1.frx":1020D
         colorIntern     =   "frmInput1.frx":10237
         colorMO         =   "frmInput1.frx":10261
         colorFocus      =   "frmInput1.frx":1028B
         colorDisabled   =   "frmInput1.frx":102B5
         colorPressed    =   "frmInput1.frx":102DF
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   1
         Left            =   1785
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":10309
         textLT          =   "frmInput1.frx":1036B
         textCT          =   "frmInput1.frx":10383
         textRT          =   "frmInput1.frx":1039B
         textLM          =   "frmInput1.frx":103B3
         textRM          =   "frmInput1.frx":103CB
         textLB          =   "frmInput1.frx":103E3
         textCB          =   "frmInput1.frx":103FB
         textRB          =   "frmInput1.frx":10413
         colorBack       =   "frmInput1.frx":1042B
         colorIntern     =   "frmInput1.frx":10455
         colorMO         =   "frmInput1.frx":1047F
         colorFocus      =   "frmInput1.frx":104A9
         colorDisabled   =   "frmInput1.frx":104D3
         colorPressed    =   "frmInput1.frx":104FD
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   4
         Left            =   1785
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   4905
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":10527
         textLT          =   "frmInput1.frx":10589
         textCT          =   "frmInput1.frx":105A1
         textRT          =   "frmInput1.frx":105B9
         textLM          =   "frmInput1.frx":105D1
         textRM          =   "frmInput1.frx":105E9
         textLB          =   "frmInput1.frx":10601
         textCB          =   "frmInput1.frx":10619
         textRB          =   "frmInput1.frx":10631
         colorBack       =   "frmInput1.frx":10649
         colorIntern     =   "frmInput1.frx":10673
         colorMO         =   "frmInput1.frx":1069D
         colorFocus      =   "frmInput1.frx":106C7
         colorDisabled   =   "frmInput1.frx":106F1
         colorPressed    =   "frmInput1.frx":1071B
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1425
         Index           =   9
         Left            =   60
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   7635
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
         textCaption     =   "frmInput1.frx":10745
         textLT          =   "frmInput1.frx":107A9
         textCT          =   "frmInput1.frx":107C1
         textRT          =   "frmInput1.frx":107D9
         textLM          =   "frmInput1.frx":107F1
         textRM          =   "frmInput1.frx":10809
         textLB          =   "frmInput1.frx":10821
         textCB          =   "frmInput1.frx":10839
         textRB          =   "frmInput1.frx":10851
         colorBack       =   "frmInput1.frx":10869
         colorIntern     =   "frmInput1.frx":10893
         colorMO         =   "frmInput1.frx":108BD
         colorFocus      =   "frmInput1.frx":108E7
         colorDisabled   =   "frmInput1.frx":10911
         colorPressed    =   "frmInput1.frx":1093B
         Orientation     =   8
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   6
         Left            =   60
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   6270
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":10965
         textLT          =   "frmInput1.frx":109C7
         textCT          =   "frmInput1.frx":109DF
         textRT          =   "frmInput1.frx":109F7
         textLM          =   "frmInput1.frx":10A0F
         textRM          =   "frmInput1.frx":10A27
         textLB          =   "frmInput1.frx":10A3F
         textCB          =   "frmInput1.frx":10A57
         textRB          =   "frmInput1.frx":10A6F
         colorBack       =   "frmInput1.frx":10A87
         colorIntern     =   "frmInput1.frx":10AB1
         colorMO         =   "frmInput1.frx":10ADB
         colorFocus      =   "frmInput1.frx":10B05
         colorDisabled   =   "frmInput1.frx":10B2F
         colorPressed    =   "frmInput1.frx":10B59
         Orientation     =   5
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   3
         Left            =   60
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   4905
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   4
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":10B83
         textLT          =   "frmInput1.frx":10BE5
         textCT          =   "frmInput1.frx":10BFD
         textRT          =   "frmInput1.frx":10C15
         textLM          =   "frmInput1.frx":10C2D
         textRM          =   "frmInput1.frx":10C45
         textLB          =   "frmInput1.frx":10C5D
         textCB          =   "frmInput1.frx":10C75
         textRB          =   "frmInput1.frx":10C8D
         colorBack       =   "frmInput1.frx":10CA5
         colorIntern     =   "frmInput1.frx":10CCF
         colorMO         =   "frmInput1.frx":10CF9
         colorFocus      =   "frmInput1.frx":10D23
         colorDisabled   =   "frmInput1.frx":10D4D
         colorPressed    =   "frmInput1.frx":10D77
         Orientation     =   5
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1365
         Index           =   0
         Left            =   60
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3540
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
         textCaption     =   "frmInput1.frx":10DA1
         textLT          =   "frmInput1.frx":10E03
         textCT          =   "frmInput1.frx":10E1B
         textRT          =   "frmInput1.frx":10E33
         textLM          =   "frmInput1.frx":10E4B
         textRM          =   "frmInput1.frx":10E63
         textLB          =   "frmInput1.frx":10E7B
         textCB          =   "frmInput1.frx":10E93
         textRB          =   "frmInput1.frx":10EAB
         colorBack       =   "frmInput1.frx":10EC3
         colorIntern     =   "frmInput1.frx":10EED
         colorMO         =   "frmInput1.frx":10F17
         colorFocus      =   "frmInput1.frx":10F41
         colorDisabled   =   "frmInput1.frx":10F6B
         colorPressed    =   "frmInput1.frx":10F95
         Orientation     =   5
         HollowFrame     =   -1  'True
         LightDirection  =   7
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   975
         Index           =   12
         Left            =   60
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   60
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
         BackColorContainer=   11983853
         SpecialEffect   =   1
         LogPixels       =   96
         SpecialEffectFactor=   3
         TextureBevelFactor=   4
         UserData        =   0.1
         textCaption     =   "frmInput1.frx":10FBF
         textLT          =   "frmInput1.frx":11037
         textCT          =   "frmInput1.frx":1104F
         textRT          =   "frmInput1.frx":11067
         textLM          =   "frmInput1.frx":1107F
         textRM          =   "frmInput1.frx":11097
         textLB          =   "frmInput1.frx":110AF
         textCB          =   "frmInput1.frx":110C7
         textRB          =   "frmInput1.frx":110DF
         colorBack       =   "frmInput1.frx":110F7
         colorIntern     =   "frmInput1.frx":11121
         colorMO         =   "frmInput1.frx":1114B
         colorFocus      =   "frmInput1.frx":11175
         colorDisabled   =   "frmInput1.frx":1119F
         colorPressed    =   "frmInput1.frx":111C9
         HollowFrame     =   -1  'True
         LightDirection  =   1
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   0
            Top             =   180
         End
      End
      Begin BTNENHLib4.BtnEnh cmdInput 
         Height          =   1605
         Index           =   13
         Left            =   3390
         TabIndex        =   61
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1500
         Width           =   1665
         _Version        =   524298
         _ExtentX        =   2937
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
         textCaption     =   "frmInput1.frx":111F3
         textLT          =   "frmInput1.frx":11257
         textCT          =   "frmInput1.frx":1126F
         textRT          =   "frmInput1.frx":11287
         textLM          =   "frmInput1.frx":1129F
         textRM          =   "frmInput1.frx":112B7
         textLB          =   "frmInput1.frx":112CF
         textCB          =   "frmInput1.frx":112E7
         textRB          =   "frmInput1.frx":112FF
         colorBack       =   "frmInput1.frx":11317
         colorIntern     =   "frmInput1.frx":11341
         colorMO         =   "frmInput1.frx":1136B
         colorFocus      =   "frmInput1.frx":11395
         colorDisabled   =   "frmInput1.frx":113BF
         colorPressed    =   "frmInput1.frx":113E9
         HollowFrame     =   -1  'True
         LightDirection  =   7
         ShapeHeadFactor =   40
         ShapeLineFactor =   40
      End
      Begin MSForms.Image Image4 
         Height          =   1185
         Left            =   1740
         Top             =   2250
         Width           =   1545
         BorderStyle     =   0
         Size            =   "2725;2090"
         VariousPropertyBits=   19
      End
      Begin MSForms.Image Image3 
         Height          =   1185
         Left            =   1740
         Top             =   1080
         Width           =   1545
         BorderStyle     =   0
         Size            =   "2725;2090"
         VariousPropertyBits=   19
      End
      Begin MSForms.Image Image1 
         Height          =   1125
         Left            =   1740
         Top             =   1140
         Width           =   1545
         BorderColor     =   255
         BackColor       =   16777215
         Size            =   "2725;1984"
         VariousPropertyBits=   19
      End
      Begin MSForms.Image Image2 
         Height          =   1125
         Left            =   1740
         Top             =   2310
         Width           =   1545
         BorderColor     =   0
         BackColor       =   16777215
         Size            =   "2725;1984"
         VariousPropertyBits=   19
      End
      Begin MSForms.TextBox txtTables 
         Height          =   705
         Left            =   2010
         TabIndex        =   65
         Top             =   1500
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
      Begin MSForms.TextBox txtCovers 
         Height          =   705
         Left            =   2010
         TabIndex        =   64
         Top             =   2640
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
      Begin MSForms.Label Label1 
         Height          =   435
         Left            =   60
         TabIndex        =   63
         Top             =   1560
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
      Begin MSForms.Label Label2 
         Height          =   435
         Left            =   60
         TabIndex        =   62
         Top             =   2670
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
   End
   Begin VB.PictureBox picSlip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9315
      Left            =   4020
      ScaleHeight     =   9285
      ScaleWidth      =   615
      TabIndex        =   34
      Top             =   1470
      Visible         =   0   'False
      Width           =   645
      Begin btButtonEx.ButtonEx cmdClose 
         Height          =   630
         Left            =   4230
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483634
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
      Begin MSForms.Label lblWorkstation 
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   8670
         Width           =   3915
         ForeColor       =   8388608
         BackColor       =   -2147483643
         VariousPropertyBits=   8388627
         Size            =   "6906;661"
         FontName        =   "Arial Narrow"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblTender 
         Height          =   375
         Left            =   4470
         TabIndex        =   40
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
         TabIndex        =   39
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
   Begin VSFlex8Ctl.VSFlexGrid grdTab 
      Height          =   7890
      Left            =   30
      TabIndex        =   9
      Top             =   1410
      Visible         =   0   'False
      Width           =   75
      _cx             =   132
      _cy             =   13917
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
      Cols            =   7
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
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   2
      Left            =   1620
      TabIndex        =   46
      Top             =   600
      Width           =   1350
      _Version        =   524298
      _ExtentX        =   2381
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Send to Table"
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
      CaptionWordWrapPerc=   88
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput1.frx":11413
      textLT          =   "frmInput1.frx":1148D
      textCT          =   "frmInput1.frx":114A5
      textRT          =   "frmInput1.frx":114BD
      textLM          =   "frmInput1.frx":114D5
      textRM          =   "frmInput1.frx":114ED
      textLB          =   "frmInput1.frx":11505
      textCB          =   "frmInput1.frx":1151D
      textRB          =   "frmInput1.frx":11535
      colorBack       =   "frmInput1.frx":1154D
      colorIntern     =   "frmInput1.frx":11577
      colorMO         =   "frmInput1.frx":115A1
      colorFocus      =   "frmInput1.frx":115CB
      colorDisabled   =   "frmInput1.frx":115F5
      colorPressed    =   "frmInput1.frx":1161F
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdLogoff 
      Height          =   1755
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1275
      _Version        =   524298
      _ExtentX        =   2249
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Exit Tabs"
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
      CaptionWordWrapPerc=   80
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput1.frx":11649
      textLT          =   "frmInput1.frx":116BB
      textCT          =   "frmInput1.frx":116D3
      textRT          =   "frmInput1.frx":116EB
      textLM          =   "frmInput1.frx":11703
      textRM          =   "frmInput1.frx":1171B
      textLB          =   "frmInput1.frx":11733
      textCB          =   "frmInput1.frx":1174B
      textRB          =   "frmInput1.frx":11763
      colorBack       =   "frmInput1.frx":1177B
      colorIntern     =   "frmInput1.frx":117A5
      colorMO         =   "frmInput1.frx":117CF
      colorFocus      =   "frmInput1.frx":117F9
      colorDisabled   =   "frmInput1.frx":11823
      colorPressed    =   "frmInput1.frx":1184D
      Orientation     =   5
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1755
      Index           =   3
      Left            =   2970
      TabIndex        =   3
      Top             =   600
      Width           =   1605
      _Version        =   524298
      _ExtentX        =   2831
      _ExtentY        =   3096
      _StockProps     =   66
      Caption         =   "Transfer Tab"
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
      textCaption     =   "frmInput1.frx":11877
      textLT          =   "frmInput1.frx":118EF
      textCT          =   "frmInput1.frx":11907
      textRT          =   "frmInput1.frx":1191F
      textLM          =   "frmInput1.frx":11937
      textRM          =   "frmInput1.frx":1194F
      textLB          =   "frmInput1.frx":11967
      textCB          =   "frmInput1.frx":1197F
      textRB          =   "frmInput1.frx":11997
      colorBack       =   "frmInput1.frx":119AF
      colorIntern     =   "frmInput1.frx":119D9
      colorMO         =   "frmInput1.frx":11A03
      colorFocus      =   "frmInput1.frx":11A2D
      colorDisabled   =   "frmInput1.frx":11A57
      colorPressed    =   "frmInput1.frx":11A81
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   945
      Index           =   6
      Left            =   4560
      TabIndex        =   6
      Top             =   1410
      Width           =   2775
      _Version        =   524298
      _ExtentX        =   4895
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
      textCaption     =   "frmInput1.frx":11AAB
      textLT          =   "frmInput1.frx":11B1F
      textCT          =   "frmInput1.frx":11B37
      textRT          =   "frmInput1.frx":11B4F
      textLM          =   "frmInput1.frx":11B67
      textRM          =   "frmInput1.frx":11B7F
      textLB          =   "frmInput1.frx":11B97
      textCB          =   "frmInput1.frx":11BAF
      textRB          =   "frmInput1.frx":11BC7
      colorBack       =   "frmInput1.frx":11BDF
      colorIntern     =   "frmInput1.frx":11C09
      colorMO         =   "frmInput1.frx":11C33
      colorFocus      =   "frmInput1.frx":11C5D
      colorDisabled   =   "frmInput1.frx":11C87
      colorPressed    =   "frmInput1.frx":11CB1
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   945
      Index           =   8
      Left            =   7320
      TabIndex        =   8
      Top             =   1410
      Width           =   2415
      _Version        =   524298
      _ExtentX        =   4260
      _ExtentY        =   1667
      _StockProps     =   66
      Caption         =   "Change Name"
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
      textCaption     =   "frmInput1.frx":11CDB
      textLT          =   "frmInput1.frx":11D51
      textCT          =   "frmInput1.frx":11D69
      textRT          =   "frmInput1.frx":11D81
      textLM          =   "frmInput1.frx":11D99
      textRM          =   "frmInput1.frx":11DB1
      textLB          =   "frmInput1.frx":11DC9
      textCB          =   "frmInput1.frx":11DE1
      textRB          =   "frmInput1.frx":11DF9
      colorBack       =   "frmInput1.frx":11E11
      colorIntern     =   "frmInput1.frx":11E3B
      colorMO         =   "frmInput1.frx":11E65
      colorFocus      =   "frmInput1.frx":11E8F
      colorDisabled   =   "frmInput1.frx":11EB9
      colorPressed    =   "frmInput1.frx":11EE3
      Orientation     =   6
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VB.Timer ScrolTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   990
      Top             =   0
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1275
      Index           =   4
      Left            =   6030
      TabIndex        =   4
      Top             =   2340
      Width           =   1860
      _Version        =   524298
      _ExtentX        =   3281
      _ExtentY        =   2249
      _StockProps     =   66
      Caption         =   "Close Tab"
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
      textCaption     =   "frmInput1.frx":11F0D
      textLT          =   "frmInput1.frx":11F7F
      textCT          =   "frmInput1.frx":11F97
      textRT          =   "frmInput1.frx":11FAF
      textLM          =   "frmInput1.frx":11FC7
      textRM          =   "frmInput1.frx":11FDF
      textLB          =   "frmInput1.frx":11FF7
      textCB          =   "frmInput1.frx":1200F
      textRB          =   "frmInput1.frx":12027
      colorBack       =   "frmInput1.frx":1203F
      colorIntern     =   "frmInput1.frx":12069
      colorMO         =   "frmInput1.frx":12093
      colorFocus      =   "frmInput1.frx":120BD
      colorDisabled   =   "frmInput1.frx":120E7
      colorPressed    =   "frmInput1.frx":12111
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
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1275
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   2340
      Width           =   1920
      _Version        =   524298
      _ExtentX        =   3387
      _ExtentY        =   2249
      _StockProps     =   66
      Caption         =   "Show All Tabs"
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
      textCaption     =   "frmInput1.frx":1213B
      textLT          =   "frmInput1.frx":121B5
      textCT          =   "frmInput1.frx":121CD
      textRT          =   "frmInput1.frx":121E5
      textLM          =   "frmInput1.frx":121FD
      textRM          =   "frmInput1.frx":12215
      textLB          =   "frmInput1.frx":1222D
      textCB          =   "frmInput1.frx":12245
      textRB          =   "frmInput1.frx":1225D
      colorBack       =   "frmInput1.frx":12275
      colorIntern     =   "frmInput1.frx":1229F
      colorMO         =   "frmInput1.frx":122C9
      colorFocus      =   "frmInput1.frx":122F3
      colorDisabled   =   "frmInput1.frx":1231D
      colorPressed    =   "frmInput1.frx":12347
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1275
      Index           =   5
      Left            =   4170
      TabIndex        =   5
      Top             =   2340
      Width           =   1860
      _Version        =   524298
      _ExtentX        =   3281
      _ExtentY        =   2249
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
      textCaption     =   "frmInput1.frx":12371
      textLT          =   "frmInput1.frx":123E5
      textCT          =   "frmInput1.frx":123FD
      textRT          =   "frmInput1.frx":12415
      textLM          =   "frmInput1.frx":1242D
      textRM          =   "frmInput1.frx":12445
      textLB          =   "frmInput1.frx":1245D
      textCB          =   "frmInput1.frx":12475
      textRB          =   "frmInput1.frx":1248D
      colorBack       =   "frmInput1.frx":124A5
      colorIntern     =   "frmInput1.frx":124CF
      colorMO         =   "frmInput1.frx":124F9
      colorFocus      =   "frmInput1.frx":12523
      colorDisabled   =   "frmInput1.frx":1254D
      colorPressed    =   "frmInput1.frx":12577
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
      TabIndex        =   1
      Top             =   1140
      Width           =   825
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1275
      Index           =   7
      Left            =   7890
      TabIndex        =   7
      Top             =   2340
      Width           =   1850
      _Version        =   524298
      _ExtentX        =   3263
      _ExtentY        =   2249
      _StockProps     =   66
      Caption         =   "View Tab"
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
      textCaption     =   "frmInput1.frx":125A1
      textLT          =   "frmInput1.frx":12611
      textCT          =   "frmInput1.frx":12629
      textRT          =   "frmInput1.frx":12641
      textLM          =   "frmInput1.frx":12659
      textRM          =   "frmInput1.frx":12671
      textLB          =   "frmInput1.frx":12689
      textCB          =   "frmInput1.frx":126A1
      textRB          =   "frmInput1.frx":126B9
      colorBack       =   "frmInput1.frx":126D1
      colorIntern     =   "frmInput1.frx":126FB
      colorMO         =   "frmInput1.frx":12725
      colorFocus      =   "frmInput1.frx":1274F
      colorDisabled   =   "frmInput1.frx":12779
      colorPressed    =   "frmInput1.frx":127A3
      Orientation     =   7
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1380
      Index           =   4
      Left            =   360
      TabIndex        =   10
      Top             =   5145
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":127CD
      textLT          =   "frmInput1.frx":127E5
      textCT          =   "frmInput1.frx":127FD
      textRT          =   "frmInput1.frx":12815
      textLM          =   "frmInput1.frx":1282D
      textRM          =   "frmInput1.frx":12845
      textLB          =   "frmInput1.frx":1285D
      textCB          =   "frmInput1.frx":12875
      textRB          =   "frmInput1.frx":1288D
      colorBack       =   "frmInput1.frx":128A5
      colorIntern     =   "frmInput1.frx":128CF
      colorMO         =   "frmInput1.frx":128F9
      colorFocus      =   "frmInput1.frx":12923
      colorDisabled   =   "frmInput1.frx":1294D
      colorPressed    =   "frmInput1.frx":12977
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1380
      Index           =   5
      Left            =   2700
      TabIndex        =   11
      Top             =   5145
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":129A1
      textLT          =   "frmInput1.frx":129B9
      textCT          =   "frmInput1.frx":129D1
      textRT          =   "frmInput1.frx":129E9
      textLM          =   "frmInput1.frx":12A01
      textRM          =   "frmInput1.frx":12A19
      textLB          =   "frmInput1.frx":12A31
      textCB          =   "frmInput1.frx":12A49
      textRB          =   "frmInput1.frx":12A61
      colorBack       =   "frmInput1.frx":12A79
      colorIntern     =   "frmInput1.frx":12AA3
      colorMO         =   "frmInput1.frx":12ACD
      colorFocus      =   "frmInput1.frx":12AF7
      colorDisabled   =   "frmInput1.frx":12B21
      colorPressed    =   "frmInput1.frx":12B4B
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1380
      Index           =   6
      Left            =   5040
      TabIndex        =   12
      Top             =   5145
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":12B75
      textLT          =   "frmInput1.frx":12B8D
      textCT          =   "frmInput1.frx":12BA5
      textRT          =   "frmInput1.frx":12BBD
      textLM          =   "frmInput1.frx":12BD5
      textRM          =   "frmInput1.frx":12BED
      textLB          =   "frmInput1.frx":12C05
      textCB          =   "frmInput1.frx":12C1D
      textRB          =   "frmInput1.frx":12C35
      colorBack       =   "frmInput1.frx":12C4D
      colorIntern     =   "frmInput1.frx":12C77
      colorMO         =   "frmInput1.frx":12CA1
      colorFocus      =   "frmInput1.frx":12CCB
      colorDisabled   =   "frmInput1.frx":12CF5
      colorPressed    =   "frmInput1.frx":12D1F
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1380
      Index           =   7
      Left            =   7440
      TabIndex        =   13
      Top             =   5145
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":12D49
      textLT          =   "frmInput1.frx":12D61
      textCT          =   "frmInput1.frx":12D79
      textRT          =   "frmInput1.frx":12D91
      textLM          =   "frmInput1.frx":12DA9
      textRM          =   "frmInput1.frx":12DC1
      textLB          =   "frmInput1.frx":12DD9
      textCB          =   "frmInput1.frx":12DF1
      textRB          =   "frmInput1.frx":12E09
      colorBack       =   "frmInput1.frx":12E21
      colorIntern     =   "frmInput1.frx":12E4B
      colorMO         =   "frmInput1.frx":12E75
      colorFocus      =   "frmInput1.frx":12E9F
      colorDisabled   =   "frmInput1.frx":12EC9
      colorPressed    =   "frmInput1.frx":12EF3
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1380
      Index           =   8
      Left            =   360
      TabIndex        =   14
      Top             =   6525
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":12F1D
      textLT          =   "frmInput1.frx":12F35
      textCT          =   "frmInput1.frx":12F4D
      textRT          =   "frmInput1.frx":12F65
      textLM          =   "frmInput1.frx":12F7D
      textRM          =   "frmInput1.frx":12F95
      textLB          =   "frmInput1.frx":12FAD
      textCB          =   "frmInput1.frx":12FC5
      textRB          =   "frmInput1.frx":12FDD
      colorBack       =   "frmInput1.frx":12FF5
      colorIntern     =   "frmInput1.frx":1301F
      colorMO         =   "frmInput1.frx":13049
      colorFocus      =   "frmInput1.frx":13073
      colorDisabled   =   "frmInput1.frx":1309D
      colorPressed    =   "frmInput1.frx":130C7
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1380
      Index           =   9
      Left            =   2700
      TabIndex        =   15
      Top             =   6525
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":130F1
      textLT          =   "frmInput1.frx":13109
      textCT          =   "frmInput1.frx":13121
      textRT          =   "frmInput1.frx":13139
      textLM          =   "frmInput1.frx":13151
      textRM          =   "frmInput1.frx":13169
      textLB          =   "frmInput1.frx":13181
      textCB          =   "frmInput1.frx":13199
      textRB          =   "frmInput1.frx":131B1
      colorBack       =   "frmInput1.frx":131C9
      colorIntern     =   "frmInput1.frx":131F3
      colorMO         =   "frmInput1.frx":1321D
      colorFocus      =   "frmInput1.frx":13247
      colorDisabled   =   "frmInput1.frx":13271
      colorPressed    =   "frmInput1.frx":1329B
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1380
      Index           =   10
      Left            =   5040
      TabIndex        =   16
      Top             =   6525
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":132C5
      textLT          =   "frmInput1.frx":132DD
      textCT          =   "frmInput1.frx":132F5
      textRT          =   "frmInput1.frx":1330D
      textLM          =   "frmInput1.frx":13325
      textRM          =   "frmInput1.frx":1333D
      textLB          =   "frmInput1.frx":13355
      textCB          =   "frmInput1.frx":1336D
      textRB          =   "frmInput1.frx":13385
      colorBack       =   "frmInput1.frx":1339D
      colorIntern     =   "frmInput1.frx":133C7
      colorMO         =   "frmInput1.frx":133F1
      colorFocus      =   "frmInput1.frx":1341B
      colorDisabled   =   "frmInput1.frx":13445
      colorPressed    =   "frmInput1.frx":1346F
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1380
      Index           =   11
      Left            =   7440
      TabIndex        =   17
      Top             =   6525
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":13499
      textLT          =   "frmInput1.frx":134B1
      textCT          =   "frmInput1.frx":134C9
      textRT          =   "frmInput1.frx":134E1
      textLM          =   "frmInput1.frx":134F9
      textRM          =   "frmInput1.frx":13511
      textLB          =   "frmInput1.frx":13529
      textCB          =   "frmInput1.frx":13541
      textRB          =   "frmInput1.frx":13559
      colorBack       =   "frmInput1.frx":13571
      colorIntern     =   "frmInput1.frx":1359B
      colorMO         =   "frmInput1.frx":135C5
      colorFocus      =   "frmInput1.frx":135EF
      colorDisabled   =   "frmInput1.frx":13619
      colorPressed    =   "frmInput1.frx":13643
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1350
      Index           =   12
      Left            =   360
      TabIndex        =   18
      Top             =   7905
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2381
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":1366D
      textLT          =   "frmInput1.frx":13685
      textCT          =   "frmInput1.frx":1369D
      textRT          =   "frmInput1.frx":136B5
      textLM          =   "frmInput1.frx":136CD
      textRM          =   "frmInput1.frx":136E5
      textLB          =   "frmInput1.frx":136FD
      textCB          =   "frmInput1.frx":13715
      textRB          =   "frmInput1.frx":1372D
      colorBack       =   "frmInput1.frx":13745
      colorIntern     =   "frmInput1.frx":1376F
      colorMO         =   "frmInput1.frx":13799
      colorFocus      =   "frmInput1.frx":137C3
      colorDisabled   =   "frmInput1.frx":137ED
      colorPressed    =   "frmInput1.frx":13817
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1350
      Index           =   13
      Left            =   2700
      TabIndex        =   19
      Top             =   7905
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2381
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":13841
      textLT          =   "frmInput1.frx":13859
      textCT          =   "frmInput1.frx":13871
      textRT          =   "frmInput1.frx":13889
      textLM          =   "frmInput1.frx":138A1
      textRM          =   "frmInput1.frx":138B9
      textLB          =   "frmInput1.frx":138D1
      textCB          =   "frmInput1.frx":138E9
      textRB          =   "frmInput1.frx":13901
      colorBack       =   "frmInput1.frx":13919
      colorIntern     =   "frmInput1.frx":13943
      colorMO         =   "frmInput1.frx":1396D
      colorFocus      =   "frmInput1.frx":13997
      colorDisabled   =   "frmInput1.frx":139C1
      colorPressed    =   "frmInput1.frx":139EB
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1350
      Index           =   14
      Left            =   5040
      TabIndex        =   20
      Top             =   7905
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2381
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":13A15
      textLT          =   "frmInput1.frx":13A2D
      textCT          =   "frmInput1.frx":13A45
      textRT          =   "frmInput1.frx":13A5D
      textLM          =   "frmInput1.frx":13A75
      textRM          =   "frmInput1.frx":13A8D
      textLB          =   "frmInput1.frx":13AA5
      textCB          =   "frmInput1.frx":13ABD
      textRB          =   "frmInput1.frx":13AD5
      colorBack       =   "frmInput1.frx":13AED
      colorIntern     =   "frmInput1.frx":13B17
      colorMO         =   "frmInput1.frx":13B41
      colorFocus      =   "frmInput1.frx":13B6B
      colorDisabled   =   "frmInput1.frx":13B95
      colorPressed    =   "frmInput1.frx":13BBF
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1350
      Index           =   15
      Left            =   7440
      TabIndex        =   21
      Top             =   7905
      Visible         =   0   'False
      Width           =   2310
      _Version        =   524298
      _ExtentX        =   4075
      _ExtentY        =   2381
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":13BE9
      textLT          =   "frmInput1.frx":13C01
      textCT          =   "frmInput1.frx":13C19
      textRT          =   "frmInput1.frx":13C31
      textLM          =   "frmInput1.frx":13C49
      textRM          =   "frmInput1.frx":13C61
      textLB          =   "frmInput1.frx":13C79
      textCB          =   "frmInput1.frx":13C91
      textRB          =   "frmInput1.frx":13CA9
      colorBack       =   "frmInput1.frx":13CC1
      colorIntern     =   "frmInput1.frx":13CEB
      colorMO         =   "frmInput1.frx":13D15
      colorFocus      =   "frmInput1.frx":13D3F
      colorDisabled   =   "frmInput1.frx":13D69
      colorPressed    =   "frmInput1.frx":13D93
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1320
      Index           =   16
      Left            =   360
      TabIndex        =   22
      Top             =   9240
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2328
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
         Size            =   9.75
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput1.frx":13DBD
      textLT          =   "frmInput1.frx":13DD5
      textCT          =   "frmInput1.frx":13DED
      textRT          =   "frmInput1.frx":13E05
      textLM          =   "frmInput1.frx":13E1D
      textRM          =   "frmInput1.frx":13E35
      textLB          =   "frmInput1.frx":13E4D
      textCB          =   "frmInput1.frx":13E65
      textRB          =   "frmInput1.frx":13E7D
      colorBack       =   "frmInput1.frx":13E95
      colorIntern     =   "frmInput1.frx":13EBF
      colorMO         =   "frmInput1.frx":13EE9
      colorFocus      =   "frmInput1.frx":13F13
      colorDisabled   =   "frmInput1.frx":13F3D
      colorPressed    =   "frmInput1.frx":13F67
      Orientation     =   8
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1320
      Index           =   17
      Left            =   2700
      TabIndex        =   23
      Top             =   9240
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2328
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":13F91
      textLT          =   "frmInput1.frx":13FA9
      textCT          =   "frmInput1.frx":13FC1
      textRT          =   "frmInput1.frx":13FD9
      textLM          =   "frmInput1.frx":13FF1
      textRM          =   "frmInput1.frx":14009
      textLB          =   "frmInput1.frx":14021
      textCB          =   "frmInput1.frx":14039
      textRB          =   "frmInput1.frx":14051
      colorBack       =   "frmInput1.frx":14069
      colorIntern     =   "frmInput1.frx":14093
      colorMO         =   "frmInput1.frx":140BD
      colorFocus      =   "frmInput1.frx":140E7
      colorDisabled   =   "frmInput1.frx":14111
      colorPressed    =   "frmInput1.frx":1413B
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1320
      Index           =   18
      Left            =   5040
      TabIndex        =   24
      Top             =   9240
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2328
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":14165
      textLT          =   "frmInput1.frx":1417D
      textCT          =   "frmInput1.frx":14195
      textRT          =   "frmInput1.frx":141AD
      textLM          =   "frmInput1.frx":141C5
      textRM          =   "frmInput1.frx":141DD
      textLB          =   "frmInput1.frx":141F5
      textCB          =   "frmInput1.frx":1420D
      textRB          =   "frmInput1.frx":14225
      colorBack       =   "frmInput1.frx":1423D
      colorIntern     =   "frmInput1.frx":14267
      colorMO         =   "frmInput1.frx":14291
      colorFocus      =   "frmInput1.frx":142BB
      colorDisabled   =   "frmInput1.frx":142E5
      colorPressed    =   "frmInput1.frx":1430F
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1320
      Index           =   19
      Left            =   7440
      TabIndex        =   25
      Top             =   9240
      Visible         =   0   'False
      Width           =   2310
      _Version        =   524298
      _ExtentX        =   4075
      _ExtentY        =   2328
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
         Size            =   9.75
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput1.frx":14339
      textLT          =   "frmInput1.frx":14351
      textCT          =   "frmInput1.frx":14369
      textRT          =   "frmInput1.frx":14381
      textLM          =   "frmInput1.frx":14399
      textRM          =   "frmInput1.frx":143B1
      textLB          =   "frmInput1.frx":143C9
      textCB          =   "frmInput1.frx":143E1
      textRB          =   "frmInput1.frx":143F9
      colorBack       =   "frmInput1.frx":14411
      colorIntern     =   "frmInput1.frx":1443B
      colorMO         =   "frmInput1.frx":14465
      colorFocus      =   "frmInput1.frx":1448F
      colorDisabled   =   "frmInput1.frx":144B9
      colorPressed    =   "frmInput1.frx":144E3
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1365
      Index           =   1
      Left            =   2700
      TabIndex        =   26
      Top             =   3795
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2408
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":1450D
      textLT          =   "frmInput1.frx":14525
      textCT          =   "frmInput1.frx":1453D
      textRT          =   "frmInput1.frx":14555
      textLM          =   "frmInput1.frx":1456D
      textRM          =   "frmInput1.frx":14585
      textLB          =   "frmInput1.frx":1459D
      textCB          =   "frmInput1.frx":145B5
      textRB          =   "frmInput1.frx":145CD
      colorBack       =   "frmInput1.frx":145E5
      colorIntern     =   "frmInput1.frx":1460F
      colorMO         =   "frmInput1.frx":14639
      colorFocus      =   "frmInput1.frx":14663
      colorDisabled   =   "frmInput1.frx":1468D
      colorPressed    =   "frmInput1.frx":146B7
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1365
      Index           =   2
      Left            =   5040
      TabIndex        =   27
      Top             =   3795
      Visible         =   0   'False
      Width           =   2430
      _Version        =   524298
      _ExtentX        =   4286
      _ExtentY        =   2408
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
         Size            =   9.75
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
      textCaption     =   "frmInput1.frx":146E1
      textLT          =   "frmInput1.frx":146F9
      textCT          =   "frmInput1.frx":14711
      textRT          =   "frmInput1.frx":14729
      textLM          =   "frmInput1.frx":14741
      textRM          =   "frmInput1.frx":14759
      textLB          =   "frmInput1.frx":14771
      textCB          =   "frmInput1.frx":14789
      textRB          =   "frmInput1.frx":147A1
      colorBack       =   "frmInput1.frx":147B9
      colorIntern     =   "frmInput1.frx":147E3
      colorMO         =   "frmInput1.frx":1480D
      colorFocus      =   "frmInput1.frx":14837
      colorDisabled   =   "frmInput1.frx":14861
      colorPressed    =   "frmInput1.frx":1488B
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1365
      Index           =   0
      Left            =   360
      TabIndex        =   28
      Top             =   3795
      Visible         =   0   'False
      Width           =   2370
      _Version        =   524298
      _ExtentX        =   4180
      _ExtentY        =   2408
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
         Size            =   9.75
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput1.frx":148B5
      textLT          =   "frmInput1.frx":148CD
      textCT          =   "frmInput1.frx":148E5
      textRT          =   "frmInput1.frx":148FD
      textLM          =   "frmInput1.frx":14915
      textRM          =   "frmInput1.frx":1492D
      textLB          =   "frmInput1.frx":14945
      textCB          =   "frmInput1.frx":1495D
      textRB          =   "frmInput1.frx":14975
      colorBack       =   "frmInput1.frx":1498D
      colorIntern     =   "frmInput1.frx":149B7
      colorMO         =   "frmInput1.frx":149E1
      colorFocus      =   "frmInput1.frx":14A0B
      colorDisabled   =   "frmInput1.frx":14A35
      colorPressed    =   "frmInput1.frx":14A5F
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdTab 
      Height          =   1370
      Index           =   3
      Left            =   7440
      TabIndex        =   29
      Top             =   3790
      Visible         =   0   'False
      Width           =   2310
      _Version        =   524298
      _ExtentX        =   4075
      _ExtentY        =   2408
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
         Size            =   9.75
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   8963553
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   3
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmInput1.frx":14A89
      textLT          =   "frmInput1.frx":14AA1
      textCT          =   "frmInput1.frx":14AB9
      textRT          =   "frmInput1.frx":14AD1
      textLM          =   "frmInput1.frx":14AE9
      textRM          =   "frmInput1.frx":14B01
      textLB          =   "frmInput1.frx":14B19
      textCB          =   "frmInput1.frx":14B31
      textRB          =   "frmInput1.frx":14B49
      colorBack       =   "frmInput1.frx":14B61
      colorIntern     =   "frmInput1.frx":14B8B
      colorMO         =   "frmInput1.frx":14BB5
      colorFocus      =   "frmInput1.frx":14BDF
      colorDisabled   =   "frmInput1.frx":14C09
      colorPressed    =   "frmInput1.frx":14C33
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin BTNENHLib4.BtnEnh cmdErr 
      Height          =   885
      Left            =   4620
      TabIndex        =   30
      Top             =   345
      Visible         =   0   'False
      Width           =   10305
      _Version        =   524298
      _ExtentX        =   18177
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
      Shape           =   1
      CornerFactor    =   15
      BackColorContainer=   12632256
      SpecialEffect   =   3
      LogPixels       =   96
      SpecialEffectFactor=   2
      UserData        =   0.1
      textCaption     =   "frmInput1.frx":14C5D
      textLT          =   "frmInput1.frx":14CE3
      textCT          =   "frmInput1.frx":14CFB
      textRT          =   "frmInput1.frx":14D13
      textLM          =   "frmInput1.frx":14D2B
      textRM          =   "frmInput1.frx":14D43
      textLB          =   "frmInput1.frx":14D5B
      textCB          =   "frmInput1.frx":14D73
      textRB          =   "frmInput1.frx":14D8B
      colorBack       =   "frmInput1.frx":14DA3
      colorIntern     =   "frmInput1.frx":14DCD
      colorMO         =   "frmInput1.frx":14DF7
      colorFocus      =   "frmInput1.frx":14E21
      colorDisabled   =   "frmInput1.frx":14E4B
      colorPressed    =   "frmInput1.frx":14E75
   End
   Begin BTNENHLib4.BtnEnh cmdFancy 
      Height          =   1275
      Index           =   0
      Left            =   2280
      TabIndex        =   44
      Top             =   2340
      Width           =   1890
      _Version        =   524298
      _ExtentX        =   3334
      _ExtentY        =   2249
      _StockProps     =   66
      Caption         =   "Change Barman"
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
      textCaption     =   "frmInput1.frx":14E9F
      textLT          =   "frmInput1.frx":14F19
      textCT          =   "frmInput1.frx":14F31
      textRT          =   "frmInput1.frx":14F49
      textLM          =   "frmInput1.frx":14F61
      textRM          =   "frmInput1.frx":14F79
      textLB          =   "frmInput1.frx":14F91
      textCB          =   "frmInput1.frx":14FA9
      textRB          =   "frmInput1.frx":14FC1
      colorBack       =   "frmInput1.frx":14FD9
      colorIntern     =   "frmInput1.frx":15003
      colorMO         =   "frmInput1.frx":1502D
      colorFocus      =   "frmInput1.frx":15057
      colorDisabled   =   "frmInput1.frx":15081
      colorPressed    =   "frmInput1.frx":150AB
      Orientation     =   8
      UseAntialias    =   0   'False
      HollowFrame     =   -1  'True
      LightDirection  =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid grdMess 
      Height          =   8265
      Left            =   10230
      TabIndex        =   45
      Top             =   1920
      Width           =   4575
      _cx             =   8070
      _cy             =   14579
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   8283198
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   8283198
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      WallPaper       =   "frmInput1.frx":150D5
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSForms.Image Image9 
      Height          =   255
      Left            =   10140
      Top             =   10200
      Width           =   4785
      SizeMode        =   1
      Size            =   "8440;450"
      Picture         =   "frmInput1.frx":16F78
   End
   Begin MSForms.Image Image8 
      Height          =   255
      Left            =   10140
      Top             =   1590
      Width           =   4755
      SizeMode        =   1
      Size            =   "8387;450"
      Picture         =   "frmInput1.frx":2944A
   End
   Begin VB.Shape Shape1 
      Height          =   8445
      Left            =   10200
      Top             =   1800
      Width           =   4635
   End
   Begin VB.Label lblServers 
      Caption         =   "Label3"
      Height          =   585
      Left            =   300
      TabIndex        =   43
      Top             =   2670
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSForms.Label lblTabs 
      Height          =   285
      Left            =   5010
      TabIndex        =   41
      Top             =   10860
      Visible         =   0   'False
      Width           =   5265
      ForeColor       =   16777215
      BackColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "You do not have any Open Tables"
      Size            =   "9287;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label lblDate 
      Height          =   285
      Left            =   540
      TabIndex        =   33
      Top             =   10860
      Width           =   4365
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7699;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblUser 
      Height          =   285
      Left            =   10410
      TabIndex        =   32
      Top             =   10860
      Width           =   4365
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "7699;503"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblKeyRegister 
      Height          =   645
      Left            =   4860
      TabIndex        =   31
      Top             =   480
      Width           =   9795
      ForeColor       =   16777215
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Size            =   "17277;1138"
      FontName        =   "Arial Narrow"
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Height          =   1815
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1545
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "2725;3201"
   End
End
Attribute VB_Name = "frmInput1"
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
    ActiveReadServer "Select * from Tab_Listing_View where User_No = " & UserRecord.User_Number
    lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Tabs"
    rs.Close
End Sub
Private Sub cmdErr_Click()
    cmdErr.Visible = False
    cmdErr.BackColor = &HF2&
    lblTransfer.Visible = False
    errTimer.Enabled = False
    If cmdErr.Caption = "Select a Table to send to or Start a New Table" Then
        picTables.Visible = False
    End If
    cmdErr.Caption = ""
    If cmdFancy(3).Enabled = False Then
        errTimer.Enabled = False
        cmdFancy(1).Enabled = True
        cmdFancy(3).Enabled = True
        cmdFancy(4).Enabled = True
        cmdFancy(5).Enabled = True
        cmdFancy(6).Enabled = True
        cmdFancy(7).Enabled = True
        cmdFancy(8).Enabled = True
        cmdLogoff.Orientation = DIR_NW
        Select Case cmdFancy(1).Caption
            Case "Show All Tabs"
                LoadTabs 0
            Case "Show Own Tabs"
                LoadTabs 1
        End Select
        lblTransfer.Tag = ""
    End If
End Sub
Private Sub cmdFancy_Click(Index As Integer)
    If errTimer.Enabled = True Then Exit Sub
    If picSlip.Visible = True Then Exit Sub
    Select Case cmdFancy(Index).Caption
        Case "Split Bill"
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to Split"
        Case "Change Barman"
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to Change Owership"
        Case "Show All Tabs"
            If UserRecord.All_Tables = True Then
                cmdFancy(Index).Caption = "Show Own Tabs"
                LoadTabs 1
            Else
                cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "You do not have Access to see all Open Tabs"
            End If
         Case "Show Own Tabs"
            If UserRecord.All_Tables = True Then
                cmdFancy(Index).Caption = "Show All Tabs"
                LoadTabs 0
            Else
                cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "You do not have Access to see all Open Tabs"
            End If
        Case "Close Tab"
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to Close"
        Case "Transfer Tab"
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to Transfer From"
        Case "Send to Table"
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to send to a Table"
        Case "Print Bill"
            If TillData.ShortTender = True Then
                cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "You have to finish Tendering to Close the Sale"
                Exit Sub
            End If
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to Print the Bill"
        Case "Split Bill"
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to Split the Bill"
        Case "View Tab"
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to View"
        Case "Change Name"
            cmdErr.Visible = True
            errTimer.Enabled = True
            cmdErr.Caption = "Select a Tab to Change the Name"
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
        Case "Cancel"
            cmdErr.Visible = False
            cmdErr.BackColor = &HF2&
            errTimer.Enabled = False
            If cmdErr.Caption = "Select a Table to send to or Start a New Table" Then
                picTables.Visible = False
            End If
            cmdErr.Caption = ""
            If cmdFancy(3).Enabled = False Then
                errTimer.Enabled = False
                cmdFancy(1).Enabled = True
                cmdFancy(3).Enabled = True
                cmdFancy(4).Enabled = True
                cmdFancy(5).Enabled = True
                cmdFancy(6).Enabled = True
                cmdFancy(7).Enabled = True
                cmdFancy(8).Enabled = True
                cmdLogoff.Orientation = DIR_NW
                Select Case cmdFancy(1).Caption
                    Case "Show All Tabs"
                        LoadTabs 0
                    Case "Show Own Tabs"
                        LoadTabs 1
                End Select
                lblTransfer.Tag = ""
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
                If Val(txtCovers.Text) < Val(lblKeyRegister.Tag) Then
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    cmdErr.Caption = "New Covers (" & txtCovers.Text & ") can't be less than the old Covers (" & Val(lblKeyRegister.Tag) & ")"
                    cmdInput(13).Tag = "2"
                    Image1.BorderColor = &H80000006
                    txtCovers.Text = ""
                    txtCovers.SetFocus
                    Exit Sub
                Else
                    ActiveUpdateServer "Update Table_Listing set Covers = " & txtCovers.Text & " where Table_No = " & txtTables.Text
                    Select Case cmdFancy(1).Caption
                        Case "Show All Tables"
                            LoadTables
                        Case "Show Own Tables"
                            LoadTables
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
                            frmSales1.cmdDept(6) = "No Sale"
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
                            frmServers.Tag = "frmInput1"
                            DoEvents
                            Screen.MousePointer = 0
                            frmServers.Show vbModal
                            Select Case lblServers.Tag
                                Case ""
                                    cmdErr.Visible = False
                                    cmdErr.BackColor = &HF2&
                                    errTimer.Enabled = False
                                    picTables.Visible = False
                                    cmdErr.Caption = ""
                                    If cmdFancy(3).Enabled = False Then
                                        errTimer.Enabled = False
                                        cmdFancy(1).Enabled = True
                                        cmdFancy(3).Enabled = True
                                        cmdFancy(4).Enabled = True
                                        cmdFancy(5).Enabled = True
                                        cmdFancy(6).Enabled = True
                                        cmdFancy(7).Enabled = True
                                        cmdFancy(8).Enabled = True
                                        cmdLogoff.Orientation = DIR_NW
                                        Select Case cmdFancy(1).Caption
                                            Case "Show All Tabs"
                                                LoadTabs 0
                                            Case "Show Own Tabs"
                                                LoadTabs 1
                                        End Select
                                        lblTransfer.Tag = ""
                                    End If
                                Case Else
                                    ActiveReadServer "Select max(User_No) as User_No,max(Doc_No) as Doc_No from Tab_Listing where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1) & " group by User_No"
                                    If rs.RecordCount > 0 Then
                                        PreviousOwner = rs.Fields("User_No")
                                        Doc_No = rs.Fields("Doc_No")
                                    Else
                                        PreviousOwner = ""
                                        Doc_No = 0
                                    End If
                                    rs.Close
                                    NewUser = lblServers.Tag
                                    ActiveUpdateServer "INSERT INTO [Table_Tansfer_Journal]([Date_Time], [Tranfering_User], [Recieving_User], [From_Table], [To_Table], [From_Tab], [To_Tab],[Invoice_No],[User_Action])" & _
                                    " VALUES(Getdate(), " & Val(PreviousOwner) & ", " & lblServers.Tag & ", Null, " & txtTables.Text & ", '" & lblTransfer.Tag & "', Null," & Doc_No & ",'Tab to New Table Transfer')"
                                    DoEvents
                                    
                                    ActiveUpdateServer "Insert Into Table_Listing " & _
                                    "SELECT " & txtTables.Text & ", " & txtCovers.Text & ", " & NewUser & ", Workstation_No, Qty, Short_Desc, Line_Total, [KeyString]," & _
                                    "[Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2]," & _
                                    "[Price_Override], [Printed], [Keyregister], [Doc_No], [Locked]," & PreviousOwner & ",Dicount_Value,Discount_Amt,Member_No  FROM [Tab_Listing]" & _
                                    " Where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
                                    ActiveUpdateServer "Delete from Tab_Listing where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
                                    cmdErr.Visible = False
                                    cmdErr.BackColor = &HF2&
                                    errTimer.Enabled = False
                                    picTables.Visible = False
                                    cmdErr.Caption = ""
                                    If cmdFancy(3).Enabled = False Then
                                        errTimer.Enabled = False
                                        cmdFancy(1).Enabled = True
                                        cmdFancy(3).Enabled = True
                                        cmdFancy(4).Enabled = True
                                        cmdFancy(5).Enabled = True
                                        cmdFancy(6).Enabled = True
                                        cmdFancy(7).Enabled = True
                                        cmdFancy(8).Enabled = True
                                        cmdLogoff.Orientation = DIR_NW
                                        Select Case cmdFancy(1).Caption
                                            Case "Show All Tabs"
                                                LoadTabs 0
                                            Case "Show Own Tabs"
                                                LoadTabs 1
                                        End Select
                                    End If
                                    lblKeyRegister = "Tab: " & Mid(lblTransfer.Tag, InStr(lblTransfer.Tag, "-") + 1) & " transfered to Table No: " & txtTables.Text
                                    lblTransfer.Tag = ""
                                    txtTables.Text = ""
                                    txtCovers.Text = ""
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
                    Screen.MousePointer = 11
                    Load frmServers
                    frmServers.Tag = "frmInput1"
                    DoEvents
                    Screen.MousePointer = 0
                    frmServers.Show vbModal
                    Select Case lblServers.Tag
                        Case ""
                            cmdErr.Visible = False
                            cmdErr.BackColor = &HF2&
                            errTimer.Enabled = False
                            picTables.Visible = False
                            cmdErr.Caption = ""
                            If cmdFancy(3).Enabled = False Then
                                errTimer.Enabled = False
                                cmdFancy(1).Enabled = True
                                cmdFancy(3).Enabled = True
                                cmdFancy(4).Enabled = True
                                cmdFancy(5).Enabled = True
                                cmdFancy(6).Enabled = True
                                cmdFancy(7).Enabled = True
                                cmdFancy(8).Enabled = True
                                cmdLogoff.Orientation = DIR_NW
                                Select Case cmdFancy(1).Caption
                                    Case "Show All Tabs"
                                        LoadTabs 0
                                    Case "Show Own Tabs"
                                        LoadTabs 1
                                End Select
                                lblTransfer.Tag = ""
                            End If
                        Case Else
                            ActiveReadServer "Select max(User_No) as User_No,Max(Doc_No) as Doc_No from Tab_Listing where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1) & " group by User_No"
                            If rs.RecordCount > 0 Then
                                PreviousOwner = rs.Fields("User_No")
                                Doc_No = rs.Fields("Doc_No")
                            Else
                                PreviousOwner = ""
                                Doc_No = 0
                            End If
                            rs.Close
                            NewUser = lblServers.Tag
                            ActiveUpdateServer "INSERT INTO [Table_Tansfer_Journal]([Date_Time], [Tranfering_User], [Recieving_User], [From_Table], [To_Table], [From_Tab], [To_Tab],[Invoice_No],[User_Action])" & _
                            " VALUES(Getdate(), " & Val(PreviousOwner) & ", " & lblServers.Tag & ", Null, " & txtTables.Text & ", '" & lblTransfer.Tag & "', Null," & Doc_No & ",'Tab to New Table Transfer')"

                            ActiveUpdateServer "Insert Into Table_Listing " & _
                            "SELECT " & txtTables.Text & ", " & txtCovers.Text & ", " & NewUser & ", Workstation_No, Qty, Short_Desc, Line_Total, [KeyString]," & _
                            "[Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2]," & _
                            "[Price_Override], [Printed], [Keyregister], [Doc_No], [Locked]," & PreviousOwner & ",0,Dicount_Value,Discount_Amt,Member_No FROM [Tab_Listing]" & _
                            " Where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
                            ActiveUpdateServer "Delete from Tab_Listing where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
                            cmdErr.Visible = False
                            cmdErr.BackColor = &HF2&
                            errTimer.Enabled = False
                            picTables.Visible = False
                            cmdErr.Caption = ""
                            If cmdFancy(3).Enabled = False Then
                                errTimer.Enabled = False
                                cmdFancy(1).Enabled = True
                                cmdFancy(3).Enabled = True
                                cmdFancy(4).Enabled = True
                                cmdFancy(5).Enabled = True
                                cmdFancy(6).Enabled = True
                                cmdFancy(7).Enabled = True
                                cmdFancy(8).Enabled = True
                                cmdLogoff.Orientation = DIR_NW
                                Select Case cmdFancy(1).Caption
                                    Case "Show All Tabs"
                                        LoadTabs 0
                                    Case "Show Own Tabs"
                                        LoadTabs 1
                                End Select
                            End If
                            lblKeyRegister = "Tab: " & Mid(lblTransfer.Tag, InStr(lblTransfer.Tag, "-") + 1) & " transfered to Table No: " & txtTables.Text
                            lblTransfer.Tag = ""
                            txtTables.Text = ""
                            txtCovers.Text = ""
                    End Select
            End Select
    End Select
    If picTables.Visible = True Then
        Select Case cmdInput(13).Tag
            Case "1": txtTables.SetFocus
            Case "2": txtCovers.SetFocus
        End Select
    End If
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
    cmdLogoff.Orientation = DIR_NW
    KeyCode = 0
    frmBar.Tag = "Show Splash"
    cmdErr.Caption = ""
    cmdErr.BackColor = &HF2&
    cmdErr.Visible = False
    ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
    DoEvents
    '*****************************
    TillData.ReturnTotal = 0
    TillData.UllageTotal = 0
    TillData.VoidTotal = 0
    TillData.Tendered = 0
    TillData.Cash = 0
    TillData.Card = 0
    TillData.Cheque = 0
    TillData.Charge = 0
    TillData.Loyalty = 0
    TillData.TaxTotal = 0
    TillData.TaxableSales = 0
    TillData.NonTaxableSales = 0
    TillData.CollectedTax = 0
    TillData.CalculatedTax = 0
    TillData.Corrects = 0
    TillData.DocNo = 0
    TillData.UserOveride = 0
    TillData.Discount = 0
    TillData.DiscountVal = 0
    TillData.TabNo = 0
    TillData.TabName = ""
    frmBar.grdMain.Rows = 1
    TillData.SaleTotal = 0
    TillData.TaxTotal = 0
    frmBar.lblCash.Caption = ""
    frmBar.lblTender.Caption = ""
    '********************************
    
    
    
    frmInput1.Hide
    Exit Sub
End Sub
Private Sub cmdTab_Click(Index As Integer)
    If picSlip.Visible = True Then Exit Sub
    DoEvents
    If picTables.Visible = False Then
        If cmdTab(Index).Picture = App.Path & "\icons\downArr.bmp" Then
            grdTab.Row = grdTab.Row + 1
            For i = 0 To 19
                If grdTab.TextMatrix(grdTab.Row, 0) = "ßArrow" Then
                    If i = 0 Then
                        cmdTab(i).Caption = ""
                        cmdTab(i).TextDescrCB.Text = ""
                        cmdTab(i).TextDescrCT.Text = ""
                        cmdTab(i).Picture = App.Path & "\icons\upArr.bmp"
                        If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                    Else
                        cmdTab(i).Caption = ""
                        cmdTab(i).TextDescrCB.Text = ""
                        cmdTab(i).TextDescrCT.Text = ""
                        cmdTab(i).Picture = App.Path & "\icons\downArr.bmp"
                        If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                        grdTab.Row = grdTab.Row - 1
                        Exit For
                    End If
                Else
                    cmdTab(i).Caption = grdTab.TextMatrix(grdTab.Row, 6)
                    cmdTab(i).Tag = grdTab.TextMatrix(grdTab.Row, 0)
                    Select Case cmdFancy(1).Caption
                        Case "Show Own Tabs"
                            cmdTab(i).TextDescrCB.OffsetY = -10
                            cmdTab(i).TextDescrCB.ColorNormal = &H800000
                            cmdTab(i).TextDescrCB.Text = grdTab.TextMatrix(grdTab.Row, 2)
                        Case "Show All Tabs"
                            cmdTab(i).TextDescrCB.Text = ""
                    End Select
                    If grdTab.TextMatrix(grdTab.Row, 3) = "True" Then
                        cmdTab(i).TextDescrCT.OffsetY = 12
                        cmdTab(i).TextDescrCT.ColorNormal = &HC0&
                        cmdTab(i).TextDescrCT.Text = "In Use"
                    Else
                        If Trim(grdTab.TextMatrix(grdTab.Row, 5)) = "" Then
                            cmdTab(i).TextDescrCB.Text = ""
                        Else
                            cmdTab(i).TextDescrCB.OffsetY = -10
                            cmdTab(i).TextDescrCB.ColorNormal = &H800000
                            cmdTab(i).TextDescrCB.Text = "From > " & grdTab.TextMatrix(grdTab.Row, 5)
                        End If
                    End If
                End If
                If grdTab.Row = grdTab.Rows - 1 Then Exit For
                grdTab.Row = grdTab.Row + 1
            Next i
            For b = i + 1 To cmdTab.Count - 1
                cmdTab(b).Caption = "1"
                cmdTab(b).Tag = ""
                cmdTab(b).Visible = False
            Next b
            Exit Sub
        End If
        If cmdTab(Index).Picture = App.Path & "\icons\upArr.bmp" Then
            cmdTab(0).Picture = ""
            While grdTab.TextMatrix(grdTab.Row, 0) <> "ßArrow"
                grdTab.Row = grdTab.Row - 1
            Wend
            grdTab.Row = grdTab.Row - 19
            For i = 0 To 19
                If grdTab.TextMatrix(grdTab.Row, 0) = "ßArrow" Then
                    If i = 0 Then
                        cmdTab(i).Caption = ""
                        cmdTab(i).TextDescrCB.Text = ""
                        cmdTab(i).TextDescrCT.Text = ""
                        cmdTab(i).Picture = App.Path & "\icons\upArr.bmp"
                        If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                    Else
                        cmdTab(i).Caption = ""
                        cmdTab(i).TextDescrCB.Text = ""
                        cmdTab(i).TextDescrCT.Text = ""
                        cmdTab(i).Picture = App.Path & "\icons\downArr.bmp"
                        If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                        grdTab.Row = grdTab.Row - 1
                        Exit For
                    End If
                Else
                    cmdTab(i).Caption = grdTab.TextMatrix(grdTab.Row, 6)
                    cmdTab(i).Tag = grdTab.TextMatrix(grdTab.Row, 0)
                    Select Case cmdFancy(1).Caption
                        Case "Show Own Tabs"
                            cmdTab(i).TextDescrCB.OffsetY = -10
                            cmdTab(i).TextDescrCB.ColorNormal = &H800000
                            cmdTab(i).TextDescrCB.Text = grdTab.TextMatrix(grdTab.Row, 2)
                        Case "Show All Tabs"
                            cmdTab(i).TextDescrCB.Text = ""
                            If Trim(grdTab.TextMatrix(grdTab.Row, 5)) = "" Then
                                cmdTab(i).TextDescrCB.Text = ""
                            Else
                                cmdTab(i).TextDescrCB.OffsetY = -10
                                cmdTab(i).TextDescrCB.ColorNormal = &H800000
                                cmdTab(i).TextDescrCB.Text = "From > " & grdTab.TextMatrix(grdTab.Row, 5)
                            End If
                    End Select
                    If grdTab.TextMatrix(grdTab.Row, 3) = "True" Then
                        cmdTab(i).TextDescrCT.OffsetY = 12
                        cmdTab(i).TextDescrCT.ColorNormal = &HC0&
                        cmdTab(i).TextDescrCT.Text = "In Use"
                    Else
                        cmdTab(i).TextDescrCT.Text = ""
                    End If
                    If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                End If
                If grdTab.Row = grdTab.Rows - 1 Then Exit For
                grdTab.Row = grdTab.Row + 1
            Next i
            For b = i + 1 To cmdTab.Count - 1
                cmdTab(b).Caption = "1"
                cmdTab(b).Tag = ""
                cmdTab(b).TextDescrCB.Text = ""
                cmdTab(b).TextDescrCT.Text = ""
                cmdTab(b).ToolTipText = ""
                cmdTab(b).Visible = False
            Next b
            Exit Sub
        End If
    Else
        If cmdTab(Index).Picture = App.Path & "\icons\downArr.bmp" Then
            grdTab.Row = grdTab.Row + 1
            For i = 0 To 19
                If grdTab.TextMatrix(grdTab.Row, 0) = "ßArrow" Then
                    If i = 0 Then
                        cmdTab(i).Caption = ""
                        cmdTab(i).TextDescrCB.Text = ""
                        cmdTab(i).TextDescrCT.Text = ""
                        cmdTab(i).Picture = App.Path & "\icons\upArr.bmp"
                        If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                    Else
                        cmdTab(i).Caption = ""
                        cmdTab(i).TextDescrCB.Text = ""
                        cmdTab(i).TextDescrCT.Text = ""
                        cmdTab(i).Picture = App.Path & "\icons\downArr.bmp"
                        If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                        grdTab.Row = grdTab.Row - 1
                        Exit For
                    End If
                Else
                    cmdTab(i).Caption = "Table No: " & grdTab.TextMatrix(grdTab.Row, 0)
                    cmdTab(i).Tag = grdTab.TextMatrix(grdTab.Row, 1)
                    cmdTab(i).TextDescrCB.OffsetY = -10
                    cmdTab(i).TextDescrCB.ColorNormal = &H800000
                    cmdTab(i).TextDescrCB.Text = grdTab.TextMatrix(grdTab.Row, 2)
                    If grdTab.TextMatrix(grdTab.Row, 3) = "True" Then
                        cmdTab(i).TextDescrCT.OffsetY = 12
                        cmdTab(i).TextDescrCT.ColorNormal = &HC0&
                        cmdTab(i).TextDescrCT.Text = "In Use"
                    End If
                End If
                If grdTab.Row = grdTab.Rows - 1 Then Exit For
                grdTab.Row = grdTab.Row + 1
            Next i
            For b = i + 1 To cmdTab.Count - 1
                cmdTab(b).Caption = "1"
                cmdTab(b).Tag = ""
                cmdTab(b).Visible = False
            Next b
            Exit Sub
        End If
        If cmdTab(Index).Picture = App.Path & "\icons\upArr.bmp" Then
            cmdTab(0).Picture = ""
            While grdTab.TextMatrix(grdTab.Row, 0) <> "ßArrow"
                grdTab.Row = grdTab.Row - 1
            Wend
            grdTab.Row = grdTab.Row - 19
            For i = 0 To 19
                If grdTab.TextMatrix(grdTab.Row, 0) = "ßArrow" Then
                    If i = 0 Then
                        cmdTab(i).Caption = ""
                        cmdTab(i).TextDescrCB.Text = ""
                        cmdTab(i).TextDescrCT.Text = ""
                        cmdTab(i).Picture = App.Path & "\icons\upArr.bmp"
                        If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                    Else
                        cmdTab(i).Caption = ""
                        cmdTab(i).TextDescrCB.Text = ""
                        cmdTab(i).TextDescrCT.Text = ""
                        cmdTab(i).Picture = App.Path & "\icons\downArr.bmp"
                        If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                        grdTab.Row = grdTab.Row - 1
                        Exit For
                    End If
                Else
                    cmdTab(i).Caption = "Table No: " & grdTab.TextMatrix(grdTab.Row, 0)
                    cmdTab(i).Tag = grdTab.TextMatrix(grdTab.Row, 1)
                    cmdTab(i).TextDescrCB.OffsetY = -10
                    cmdTab(i).TextDescrCB.ColorNormal = &H800000
                    cmdTab(i).TextDescrCB.Text = grdTab.TextMatrix(grdTab.Row, 2)
                    If grdTab.TextMatrix(grdTab.Row, 3) = "True" Then
                        cmdTab(i).TextDescrCT.OffsetY = 12
                        cmdTab(i).TextDescrCT.ColorNormal = &HC0&
                        cmdTab(i).TextDescrCT.Text = "In Use"
                    Else
                        cmdTab(i).TextDescrCT.Text = ""
                    End If
                    If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
                End If
                If grdTab.Row = grdTab.Rows - 1 Then Exit For
                grdTab.Row = grdTab.Row + 1
            Next i
            For b = i + 1 To cmdTab.Count - 1
                cmdTab(b).Caption = "1"
                cmdTab(b).Tag = ""
                cmdTab(b).TextDescrCB.Text = ""
                cmdTab(b).TextDescrCT.Text = ""
                cmdTab(b).ToolTipText = ""
                cmdTab(b).Visible = False
            Next b
            Exit Sub
        End If
    End If
    If cmdTab(Index).Caption = "" Then
        If Val(cmdTab(Index).Tag) = 0 Then
            ActiveReadServer1 "Select isnull(Max(Tab_No),0)+1 as Tab_No from Tab_Listing as Tab_No"
            cmdTab(Index).Tag = rs1.Fields("Tab_No")
            rs1.Close
            ActiveUpdateServer "Update Tab_Listing set Tab_No = " & cmdTab(Index).Tag & " where Tab_No=0"
        End If
    End If
    If errTimer.Enabled = True Then
        Select Case cmdErr.Caption
            Case "Select a Tab to Split"
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                If Val(cmdTab(Index).Tag) <> Int(Val(cmdTab(Index).Tag)) Then
                    cmdErr.Visible = True
                    errTimer.Enabled = True
                    cmdErr.Caption = "Invalid Key Pressed"
                    Exit Sub
                End If
                
                ActiveReadServer "Select * from Tab_Listing_view where Tab_No = " & Val(cmdTab(Index).Tag)
                If Workstation_No <> rs.Fields("Workstation_No") Then
                    If rs.Fields("Locked") = True Then
                        cmdErr.Visible = True
                        errTimer.Enabled = True
                        cmdErr.Caption = "This Tab is already in opened by another User"
                        cmdTab(Index).TextDescrCT.OffsetY = 12
                        cmdTab(Index).TextDescrCT.ColorNormal = &HC0&
                        cmdTab(Index).TextDescrCT.Text = "In Use"
                        rs.Close
                        Exit Sub
                    End If
                End If
                rs.Close
                TillData.TabNo = Val(cmdTab(Index).Tag)
                frmBar.LoadOldTab Val(cmdTab(Index).Tag)
                frmBar.picSlip.Visible = True
                frmBar.cmdSlip.Caption = "Close Slip"
                frmBar.cmdFancy(3).Caption = "Add to Tab"
                Panel_no = 2
                Key_Function ("Split Bill")
            Case "Select a Tab to Print the Bill"
                Panel_no = 2
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                picSlip.Visible = False
                ActiveReadServer "Select * from Tab_Listing_view where Tab_No = " & cmdTab(Index).Tag
                If Workstation_No <> rs.Fields("Workstation_No") Then
                    If rs.Fields("Locked") = True Then
                        cmdErr.Visible = True
                        errTimer.Enabled = True
                        cmdErr.Caption = "This Tab is already in opened by another User"
                        cmdTab(Index).TextDescrCT.OffsetY = 12
                        cmdTab(Index).TextDescrCT.ColorNormal = &HC0&
                        cmdTab(Index).TextDescrCT.Text = "In Use"
                        rs.Close
                        Exit Sub
                    End If
                End If
                rs.Close
                TillData.TabNo = cmdTab(Index).Tag
                frmBar.LoadOldTab cmdTab(Index).Tag
                DoEvents
                '************** Kotie 20-03-2013 07:35
                If TillData.Print_Count > 0 Then
                    If UserRecord.Reprint = False Then
                        TillData.UserOveride = 0
                        Load frmValidate
                        frmValidate.Tag = "Reprint"
                        frmValidate.Show vbModal
                        If frmValidate.Tag = "0" Or frmValidate.Tag = "" Then
                            ActiveUpdateServer "Update Table_Listing set Locked =0 where Table_No= " & TillData.TableNo
                            errTimer = True
                            cmdErr.Caption = "Not allowed"
                            cmdErr.Visible = True
                            Exit Sub
                        Else
                      
                        End If
                    End If
                End If
                '********************
               ' frmBar.LoadOldTab cmdTab(Index).Tag
               ActiveUpdateServer ("Insert into Print_Journal (User_No,doc_No,Doc_Type,DateTimePrinted, User_override, Table_no)VALUES(" & UserRecord.User_Number & "," & TillData.DocNo & ",'Bill Print', getdate(), '" & TillData.UserOveride & "','" & TillData.TabNo & "')")
                PrintSlip "Print Bill Tab"
                'cmdFancy_Click (6)
                With frmBar
                    ActiveUpdateServer "Update Tab_Listing set Locked =0 where Tab_No= " & TillData.TabNo
                    .lblTab = ""
                    .grdMain.Rows = 1
                    .lblCash.Caption = ""
                    .lblTender.Caption = "0.00"
                    TillData.DocNo = 0
                    TillData.TabNo = 0
                    TillData.TabName = ""
                    GlobalMode = TillMode.FinMode
                    DoEvents
                    lblKeyRegister = "Printed Bill for Tab: " & cmdTab(Index).Caption
                    Exit Sub
                End With
            Case "Select a Table to send to or Start a New Table"
                ActiveReadServer "Select max(User_No) as User_No,Max(Doc_No) as Doc_No from Tab_Listing where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1) & " group by User_No"
                If rs.RecordCount > 0 Then
                    PreviousOwner = rs.Fields("User_No")
                    Doc_No = rs.Fields("Doc_No")
                Else
                    PreviousOwner = ""
                    Doc_No = 0
                End If
                rs.Close
                For i = 0 To grdTab.Rows - 1
                    If Val(Mid(cmdTab(Index).Caption, InStr(cmdTab(Index).Caption, ":") + 1)) = Val(grdTab.TextMatrix(i, 0)) Then
                        NewUser = Val(Mid(grdTab.TextMatrix(i, 2), 1, InStr(grdTab.TextMatrix(i, 2), "-") - 1))
                        Exit For
                    End If
                Next i
                ActiveUpdateServer "INSERT INTO [Table_Tansfer_Journal]([Date_Time], [Tranfering_User], [Recieving_User], [From_Table], [To_Table], [From_Tab], [To_Tab],[Invoice_No],[User_Action])" & _
                " VALUES(Getdate(), " & Val(PreviousOwner) & ", " & NewUser & ", Null, " & Val(Mid(cmdTab(Index).Caption, InStr(cmdTab(Index).Caption, ":") + 1)) & ", '" & lblTransfer.Tag & "', Null," & Val(Doc_No) & ",'Tab to Table Transfer')"

                DoEvents
                ActiveUpdateServer "Insert Into Table_Listing " & _
                "SELECT " & Val(Mid(cmdTab(Index).Caption, InStr(cmdTab(Index).Caption, ":") + 1)) & ", " & cmdTab(Index).Tag & ", " & NewUser & ", " & Workstation_No & ", Qty, Short_Desc, Line_Total, [KeyString]," & _
                "[Cost], [Tax_Rate], [Tax_Type], [Extra_Function], [Product_Code], [Dept_No], [Kitchen1], [Kitchen2]," & _
                "[Price_Override], [Printed], [Keyregister], [Doc_No], [Locked]," & PreviousOwner & ",User_Overide, Dicount_Value,Discount_Amt,Member_no FROM [Tab_Listing]" & _
                " Where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
                ActiveUpdateServer "Delete from Tab_Listing where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
                ActiveUpdateServer "Update Table_Listing set Previous_Owner = '" & PreviousOwner & "',Workstation_No = " & Workstation_No & " where Table_No = " & Val(Mid(cmdTab(Index).Caption, InStr(cmdTab(Index).Caption, ":") + 1))
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                picTables.Visible = False
                cmdErr.Caption = ""
                TableNo = Val(Mid(cmdTab(Index).Caption, InStr(cmdTab(Index).Caption, ":") + 1))
                If cmdFancy(3).Enabled = False Then
                    errTimer.Enabled = False
                    cmdFancy(1).Enabled = True
                    cmdFancy(3).Enabled = True
                    cmdFancy(4).Enabled = True
                    cmdFancy(5).Enabled = True
                    cmdFancy(6).Enabled = True
                    cmdFancy(7).Enabled = True
                    cmdFancy(8).Enabled = True
                    cmdLogoff.Orientation = DIR_NW
                    Select Case cmdFancy(1).Caption
                        Case "Show All Tabs"
                            LoadTabs 0
                        Case "Show Own Tabs"
                            LoadTabs 1
                    End Select
                End If
                lblKeyRegister = "Tab: " & Mid(lblTransfer.Tag, InStr(lblTransfer.Tag, "-") + 1) & " transfered to Table No: " & TableNo
                lblTransfer.Tag = ""
                txtTables.Text = ""
                txtCovers.Text = ""
            Case "Select a Tab to Change Owership"
                Screen.MousePointer = 11
                Load frmServers
                frmServers.Tag = "frmInput1"
                DoEvents
                Screen.MousePointer = 0
                frmServers.Show vbModal
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                If lblServers.Tag <> "" Then
                    ActiveReadServer "Select max(User_No) as User_No,Max(Doc_No) as Doc_No from Tab_Listing where Tab_No = " & cmdTab(Index).Tag & " group by User_No"
                    If rs.RecordCount > 0 Then
                        PreviousOwner = rs.Fields("User_No")
                        Doc_No = rs.Fields("Doc_No")
                    Else
                        Doc_No = 0
                        PreviousOwner = ""
                    End If
                    rs.Close
                
                    ActiveUpdateServer "INSERT INTO [Table_Tansfer_Journal]([Date_Time], [Tranfering_User], [Recieving_User], [From_Table], [To_Table], [From_Tab], [To_Tab],[Invoice_No],[User_Action])" & _
                    " VALUES(Getdate(), " & Val(PreviousOwner) & ", " & lblServers.Tag & ", Null, Null, '" & cmdTab(Index).Tag & "-" & cmdTab(Index).Caption & "','" & cmdTab(Index).Tag & "-" & cmdTab(Index).Caption & "'," & Doc_No & ",'Barman Change')"
                    ActiveUpdateServer "Update Tab_Listing set Previous_Owner = User_No, User_No = " & lblServers.Tag & " where Tab_No = " & cmdTab(Index).Tag
                End If
                Select Case cmdFancy(1).Caption
                    Case "Show All Tabs"
                        LoadTabs 0
                    Case "Show Own Tabs"
                        LoadTabs 1
                End Select
            Case "Select a Tab to send to a Table"
                cmdFancy(1).Enabled = False
                cmdFancy(3).Enabled = False
                cmdFancy(4).Enabled = False
                cmdFancy(5).Enabled = False
                cmdFancy(6).Enabled = False
                cmdFancy(7).Enabled = False
                cmdFancy(8).Enabled = False
                cmdLogoff.Orientation = DIR_WEST
                cmdErr.Caption = "Select a Table to send to or Start a New Table"
                lblTransfer.Caption = "Transfering Tab: " & cmdTab(Index).Caption
                lblTransfer.Tag = Val(cmdTab(Index).Tag) & "-" & cmdTab(Index).Caption
                txtTables.Text = ""
                txtCovers.Text = ""
                picTables.Visible = True
                txtTables.SetFocus
                LoadTables
                Exit Sub
            Case "Select a Tab to Transfer to"
                lblTransfer.Visible = False
                ActiveReadServer "Select max(User_No) as User_No,Max(Doc_No) as Doc_No from Tab_Listing where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1) & " group by User_No"
                If rs.RecordCount > 0 Then
                    PreviousOwner = rs.Fields("User_No")
                    Doc_No = rs.Fields("Doc_No")
                Else
                    Doc_No = 0
                    PreviousOwner = ""
                End If
                rs.Close
                For i = 0 To grdTab.Rows - 1
                    If Val(cmdTab(Index).Tag) = Val(grdTab.TextMatrix(i, 0)) Then
                        NewUser = Val(Mid(grdTab.TextMatrix(i, 2), 1, InStr(grdTab.TextMatrix(i, 2), "-") - 1))
                        Exit For
                    End If
                Next i
                ActiveUpdateServer "INSERT INTO [Table_Tansfer_Journal]([Date_Time], [Tranfering_User], [Recieving_User], [From_Table], [To_Table], [From_Tab], [To_Tab],[Invoice_No],[User_Action])" & _
                " VALUES(Getdate(), " & Val(PreviousOwner) & ", " & NewUser & ", Null, Null, '" & lblTransfer.Tag & "','" & cmdTab(Index).Tag & "-" & cmdTab(Index).Caption & "'," & Doc_No & ",'Tab to Tab Transfer')"
                DoEvents
                ActiveUpdateServer "Update Tab_Listing set User_No = " & NewUser & ",Tab_no = " & Val(cmdTab(Index).Tag) & ",Tab_Name = '" & cmdTab(Index).Caption & "' where Tab_No = " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
                ActiveUpdateServer "Update Tab_Listing set Locked = 0,Previous_Owner = '" & PreviousOwner & "' where Tab_No = " & cmdTab(Index).Tag
                cmdFancy(1).Enabled = True
                cmdFancy(3).Enabled = True
                cmdFancy(4).Enabled = True
                cmdFancy(5).Enabled = True
                cmdFancy(6).Enabled = True
                cmdFancy(7).Enabled = True
                cmdFancy(8).Enabled = True
                cmdLogoff.Orientation = DIR_NW
                Select Case cmdFancy(1).Caption
                    Case "Show All Tabs"
                        LoadTabs 0
                    Case "Show Own Tabs"
                        LoadTabs 1
                End Select
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                lblKeyRegister = "Tab No: " & Mid(lblTransfer.Tag, InStr(lblTransfer.Tag, "-") + 1) & " transfered to Tab: " & cmdTab(Index).Caption
                lblTransfer.Tag = ""
                Exit Sub
            Case "Select a Tab to Transfer From"
                cmdFancy(1).Enabled = False
                cmdFancy(3).Enabled = False
                cmdFancy(4).Enabled = False
                cmdFancy(5).Enabled = False
                cmdFancy(6).Enabled = False
                cmdFancy(7).Enabled = False
                cmdFancy(8).Enabled = False
                cmdLogoff.Orientation = DIR_WEST
                cmdErr.Caption = "Select a Tab to Transfer to"
                lblTransfer.Visible = True
                lblTransfer.Caption = "Transfering Tab: " & cmdTab(Index).Caption
                lblTransfer.Tag = Val(cmdTab(Index).Tag) & "-" & cmdTab(Index).Caption
                Select Case cmdFancy(1).Caption
                    Case "Show All Tabs"
                        LoadTabs 2
                    Case "Show Own Tabs"
                        LoadTabs 3
                End Select
                Exit Sub
            Case "Select a Tab to Change the Name"
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                frmBar.lblKeyRegister = "Please Enter a new Tab Name"
                Screen.MousePointer = 11
                Load frmKeyBoard
                frmKeyBoard.Tag = "Tabs"
                DoEvents
                frmBar.Tag = "1"
                frmKeyBoard.Show vbModal
                Select Case frmBar.lblKeyRegister.Caption
                    Case ""
                        Exit Sub
                    Case Else
                        ActiveUpdateServer "Update Tab_Listing set Tab_Name = '" & frmBar.lblKeyRegister.Caption & "' where Tab_No =" & cmdTab(Index).Tag
                End Select
                Select Case cmdFancy(1).Caption
                    Case "Show All Tabs"
                        LoadTabs 0
                    Case "Show Own Tabs"
                        LoadTabs 1
                End Select
            Case "Select a Tab to Close"
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                ActiveReadServer "Select * from Tab_Listing_view where Tab_No = " & Val(cmdTab(Index).Tag)
                If Workstation_No <> rs.Fields("Workstation_No") Then
                    If rs.Fields("Locked") = True Then
                        cmdErr.Visible = True
                        errTimer.Enabled = True
                        cmdErr.Caption = "This Tab is already in opened by another User"
                        cmdTab(Index).TextDescrCT.OffsetY = 12
                        cmdTab(Index).TextDescrCT.ColorNormal = &HC0&
                        cmdTab(Index).TextDescrCT.Text = "In Use"
                        rs.Close
                        Exit Sub
                    End If
                End If
                rs.Close
                TillData.TabNo = Val(cmdTab(Index).Tag)
                frmBar.LoadOldTab Val(cmdTab(Index).Tag)
                Panel_no = 2
                Key_Function ("Close Tab")
            Case "Select a Tab to View"
                cmdErr.Visible = False
                cmdErr.BackColor = &HF2&
                errTimer.Enabled = False
                picSlip.Visible = True
                lblKeyRegister = "Viewing Tab: " & cmdTab(Index).Caption
                LoadTab cmdTab(Index).Tag
                grdMain.SetFocus
                Exit Sub
        End Select
    Else
        If picSlip.Visible = True Then Exit Sub
        If cmdTab(Index).TextDescrCT.Text = "In Use" Then
            cmdErr.Visible = True
            errTimer.Enabled = True
            Select Case cmdFancy(1).Caption
                Case "Show Own Tabs"
                    If UserRecord.User_Number <> Val(Trim(Mid(cmdTab(Index).TextDescrCB.Text, 1, InStr(cmdTab(Index).TextDescrCB.Text, "-") - 1))) Then
                        cmdErr.Caption = "This Tab is already in opened by another User"
                    Else
                        cmdErr.Caption = "This Tab is already open on another Workstation"
                    End If
                Case "Show All Tabs"
                    cmdErr.Caption = "This Tab is already open on another Workstation"
            End Select
            Exit Sub
        End If
        ActiveReadServer "Select * from Tab_Listing_view where Tab_No = " & cmdTab(Index).Tag
        If Workstation_No <> rs.Fields("Workstation_No") Then
            If rs.Fields("Locked") = True Then
                cmdErr.Visible = True
                errTimer.Enabled = True
                cmdErr.Caption = "This Tab is already in opened by another User"
                cmdTab(Index).TextDescrCT.OffsetY = 12
                cmdTab(Index).TextDescrCT.ColorNormal = &HC0&
                cmdTab(Index).TextDescrCT.Text = "In Use"
                rs.Close
                Exit Sub
            End If
        End If
        rs.Close
        TillData.TabNo = cmdTab(Index).Tag
        frmBar.LoadOldTab cmdTab(Index).Tag
        frmBar.picSlip.Visible = True
        frmBar.cmdSlip.Caption = "Close Slip"
        frmBar.cmdFancy(3).Caption = "Add to Tab"
    End If
End Sub
Private Sub LoadTab(Tab_Number)
    ActiveReadServer "Select * from Tab_Listing where Tab_No= " & Tab_Number & " order by Line_No"
    grdMain.Rows = 1
    grdMain.ColHidden(14) = True
    lblWorkstation.Caption = "Opened on Workstation No: " & rs.Fields("Workstation_No")
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
    If Me.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            Me.Controls(i).Width = Me.Controls(i).Width * 0.782
            Me.Controls(i).Left = Me.Controls(i).Left * 0.782
            Me.Controls(i).Height = Me.Controls(i).Height * 0.79
            Me.Controls(i).top = Me.Controls(i).top * 0.79
            Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.79)
            Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.78)
        Next i
        On Error GoTo 0
        newBack.Width = Me.Width
        newBack.Height = Me.Height
    End If
    Screen.MousePointer = 0
    If cmdFancy(3).Enabled = True Then
        lblDate.Caption = Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS")
        lblUser.Caption = Trim(UserRecord.FirstName) & " " & Trim(UserRecord.LastName)
        cmdErr.Caption = ""
        cmdFancy(1).Caption = "Show All Tabs"
        LoadTabs 0
        lblKeyRegister.Tag = ""
        picHoldFocus.SetFocus
    End If
    ActiveReadServer "Select * from Notice_Board order by Line_No"
    grdMess.Row = 0
    While Not rs.EOF
        grdMess.Row = grdMess.Row + 1
        grdMess.TextMatrix(grdMess.Row, 1) = rs.Fields("Description") & ""
        Select Case rs.Fields("Style") & ""
            Case "Sub Header"
                grdMess.TextMatrix(grdMess.Row, 0) = ">"
                grdMess.TextMatrix(grdMess.Row, 0) = ">"
                grdMess.Select grdMess.Row, 0, grdMess.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
            Case "Sub Header Underline"
                grdMess.TextMatrix(grdMess.Row, 0) = ">"
                grdMess.TextMatrix(grdMess.Row, 0) = ">"
                grdMess.Select grdMess.Row, 0, grdMess.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
            Case "Normal"
                grdMess.TextMatrix(grdMess.Row, 0) = ""
                grdMess.TextMatrix(grdMess.Row, 0) = ""
                grdMess.Select grdMess.Row, 0, grdMess.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 0, 0, 0
            Case "Normal Underline"
                grdMess.TextMatrix(grdMess.Row, 0) = ""
                grdMess.TextMatrix(grdMess.Row, 0) = ""
                grdMess.Select grdMess.Row, 0, grdMess.Row, 1
                grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
        End Select
        rs.MoveNext
    Wend
    rs.Close
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
    grdMess.Col = 0
    grdMess.Row = 0
    grdMess.ColWidth(0) = 180
    grdMess.Rows = grdMess.Rows + 1
    grdMess.Row = grdMess.Rows - 1
    grdMess.TextMatrix(0, 0) = "Daily Notice Board"
    grdMess.TextMatrix(0, 1) = "Daily Notice Board"
    grdMess.MergeRow(0) = True
    grdMess.CellAlignment = flexAlignCenterCenter
    grdMess.Select 0, 0, grdrow, 1
    grdMess.CellBorder vbWhite, 0, 0, 0, 1, 0, 1
    grdMess.CellFontBold = True
End Sub
Private Sub LoadTabs(Action)
    grdTab.Rows = 0
    cmdTab(0).Caption = ""
    cmdTab(0).Picture = ""
    DoEvents
    Select Case Action
        Case 0
            ActiveReadServer "Select * from Tab_Listing_View where User_No = " & UserRecord.User_Number
        Case 1
            ActiveReadServer "Select * from Tab_Listing_View"
        Case 2
            ActiveReadServer "Select * from Tab_Listing_View where User_No = " & UserRecord.User_Number & " and Tab_No <> " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
        Case 3
            ActiveReadServer "Select * from Tab_Listing_View where Tab_No <> " & Mid(lblTransfer.Tag, 1, InStr(lblTransfer.Tag, "-") - 1)
    End Select
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdTab.Rows = grdTab.Rows + 1
        If i < 19 And Not rs.EOF Then
            cmdTab(i).Caption = rs.Fields("Tab_Name") & ""
            cmdTab(i).Tag = rs.Fields("Tab_No")
            If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
            grdTab.Row = grdTab.Rows - 1
            grdTab.TextMatrix(grdTab.Rows - 1, 0) = rs.Fields("Tab_No")
            grdTab.TextMatrix(grdTab.Rows - 1, 1) = rs.Fields("Covers")
            grdTab.TextMatrix(grdTab.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTab.TextMatrix(grdTab.Rows - 1, 3) = rs.Fields("Locked")
            grdTab.TextMatrix(grdTab.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTab.TextMatrix(grdTab.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            grdTab.TextMatrix(grdTab.Rows - 1, 6) = rs.Fields("Tab_Name") & ""
            If Action = 1 Or Action = 3 Then
                cmdTab(i).TextDescrCB.OffsetY = -10
                cmdTab(i).TextDescrCB.ColorNormal = &H800000
                cmdTab(i).TextDescrCB.Text = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
                If rs.Fields("Locked") = True Then
                    cmdTab(i).TextDescrCT.OffsetY = 12
                    cmdTab(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTab(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTab(i).TextDescrCT.Text = ""
                End If
            Else
                cmdTab(i).TextDescrCB.Text = ""
                If rs.Fields("Previous_Owner") = rs.Fields("User_No") Then
                    cmdTab(i).TextDescrCB.Text = ""
                Else
                    cmdTab(i).TextDescrCB.OffsetY = -10
                    cmdTab(i).TextDescrCB.ColorNormal = &H800000
                    If rs.Fields("Previous_Name") & "" <> "" Then
                        If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                            cmdTab(i).TextDescrCB.Text = "From > " & rs.Fields("Previous_Name") & ""
                        End If
                    End If
                End If
                If rs.Fields("Locked") = True And rs.Fields("Workstation_No") <> Workstation_No Then
                    cmdTab(i).TextDescrCT.OffsetY = 12
                    cmdTab(i).TextDescrCT.ColorNormal = &HC0&
                    cmdTab(i).TextDescrCT.Text = "In Use"
                Else
                    cmdTab(i).TextDescrCT.Text = ""
                End If
            End If
        Else
            If b = 0 Then
                grdTab.TextMatrix(grdTab.Rows - 1, 0) = "ßArrow"
                grdTab.Rows = grdTab.Rows + 1
                If i = 19 Then
                    cmdTab(19).Caption = ""
                    cmdTab(19).Picture = App.Path & "\icons\downArr.bmp"
                    cmdTab(i).TextDescrCB.Text = ""
                    cmdTab(i).TextDescrCT.Text = ""
                    If cmdTab(19).Visible = False Then cmdTab(19).Visible = True
                End If
            End If
            b = b + 1
            grdTab.TextMatrix(grdTab.Rows - 1, 0) = rs.Fields("Tab_No")
            grdTab.TextMatrix(grdTab.Rows - 1, 1) = rs.Fields("Covers")
            grdTab.TextMatrix(grdTab.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTab.TextMatrix(grdTab.Rows - 1, 3) = rs.Fields("Locked")
            grdTab.TextMatrix(grdTab.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTab.TextMatrix(grdTab.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            grdTab.TextMatrix(grdTab.Rows - 1, 6) = rs.Fields("Tab_Name")
            If b = 18 Then b = 0
        End If
        rs.MoveNext
    Wend
    Select Case Action
        Case 0, 2
            If rs.RecordCount = 1 Then
                lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Tab"
            Else
                lblKeyRegister.Caption = "You have " & rs.RecordCount & " Open Tabs"
            End If
        Case 1, 3
            If rs.RecordCount = 1 Then
                lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Tab"
            Else
                lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Tabs"
            End If
    End Select
    If rs.RecordCount > 0 Then
        lblTabs.Visible = False
    Else
        lblTabs.Visible = True
    End If
    rs.Close
    For b = i + 1 To cmdTab.Count - 1
       cmdTab(b).Caption = "0"
       cmdTab(b).Visible = False
    Next b
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
Private Sub LoadTables()
    grdTab.Rows = 0
    cmdTab(0).Caption = ""
    cmdTab(0).Picture = ""
    DoEvents
    ActiveReadServer "Select * from Table_Listing_View"
    i = -1
    b = 0
    While Not rs.EOF
        i = i + 1
        grdTab.Rows = grdTab.Rows + 1
        If i < 19 And Not rs.EOF Then
            cmdTab(i).Caption = "Table No: " & rs.Fields("Table_No")
            cmdTab(i).Tag = rs.Fields("Covers")
            If cmdTab(i).Visible = False Then cmdTab(i).Visible = True
            grdTab.Row = grdTab.Rows - 1
            grdTab.TextMatrix(grdTab.Rows - 1, 0) = rs.Fields("Table_No")
            grdTab.TextMatrix(grdTab.Rows - 1, 1) = rs.Fields("Covers")
            grdTab.TextMatrix(grdTab.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTab.TextMatrix(grdTab.Rows - 1, 3) = rs.Fields("Locked")
            grdTab.TextMatrix(grdTab.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTab.TextMatrix(grdTab.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            cmdTab(i).TextDescrCB.OffsetY = -10
            cmdTab(i).TextDescrCB.ColorNormal = &H800000
            cmdTab(i).TextDescrCB.Text = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            If rs.Fields("Locked") = True Then
                cmdTab(i).TextDescrCT.OffsetY = 12
                cmdTab(i).TextDescrCT.ColorNormal = &HC0&
                cmdTab(i).TextDescrCT.Text = "In Use"
            Else
                cmdTab(i).TextDescrCT.Text = ""
            End If
        Else
            If b = 0 Then
                grdTab.TextMatrix(grdTab.Rows - 1, 0) = "ßArrow"
                grdTab.Rows = grdTab.Rows + 1
                If i = 19 Then
                    cmdTab(19).Caption = ""
                    cmdTab(19).Picture = App.Path & "\icons\downArr.bmp"
                    cmdTab(i).TextDescrCB.Text = ""
                    cmdTab(i).TextDescrCT.Text = ""
                    If cmdTab(19).Visible = False Then cmdTab(19).Visible = True
                End If
            End If
            b = b + 1
            grdTab.TextMatrix(grdTab.Rows - 1, 0) = rs.Fields("Table_No")
            grdTab.TextMatrix(grdTab.Rows - 1, 1) = rs.Fields("Covers")
            grdTab.TextMatrix(grdTab.Rows - 1, 2) = rs.Fields("User_No") & " - " & rs.Fields("User_Name")
            grdTab.TextMatrix(grdTab.Rows - 1, 3) = rs.Fields("Locked")
            grdTab.TextMatrix(grdTab.Rows - 1, 4) = rs.Fields("Workstation_No")
            If Val(UserRecord.User_Number) <> Val(rs.Fields("Previous_Owner") & "") Then
                grdTab.TextMatrix(grdTab.Rows - 1, 5) = rs.Fields("Previous_Name") & ""
            End If
            If b = 18 Then b = 0
        End If
        rs.MoveNext
    Wend
    If rs.RecordCount = 1 Then
        lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Table"
    Else
        lblKeyRegister.Caption = "There is " & rs.RecordCount & " Open Tables"
    End If
    If rs.RecordCount > 0 Then
        lblTabs.Visible = False
    Else
        lblTabs.Visible = True
    End If
    rs.Close
    For b = i + 1 To cmdTab.Count - 1
       cmdTab(b).Caption = "0"
       cmdTab(b).Visible = False
    Next b
End Sub

