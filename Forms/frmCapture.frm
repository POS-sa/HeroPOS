VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{40B5CE80-C5A8-11D2-8183-00002440DFD8}#8.10#0"; "3dabm8u.ocx"
Begin VB.Form frmCapture 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   10890
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCapture.frx":0000
   ScaleHeight     =   10890
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFloat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00AAD7EC&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6810
      ScaleHeight     =   585
      ScaleWidth      =   2280
      TabIndex        =   68
      Top             =   2100
      Visible         =   0   'False
      Width           =   2315
      Begin VB.TextBox txtFloat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00AAD7EC&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -15
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   60
         Width           =   2205
      End
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1395
      Index           =   0
      Left            =   4650
      TabIndex        =   26
      Top             =   3780
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2461
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
      CornerFactor    =   15
      Surface         =   1
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCapture.frx":8768
      textLT          =   "frmCapture.frx":87CA
      textCT          =   "frmCapture.frx":87E2
      textRT          =   "frmCapture.frx":87FA
      textLM          =   "frmCapture.frx":8812
      textRM          =   "frmCapture.frx":882A
      textLB          =   "frmCapture.frx":8842
      textCB          =   "frmCapture.frx":885A
      textRB          =   "frmCapture.frx":8872
      colorBack       =   "frmCapture.frx":888A
      colorIntern     =   "frmCapture.frx":88B4
      colorMO         =   "frmCapture.frx":88DE
      colorFocus      =   "frmCapture.frx":8908
      colorDisabled   =   "frmCapture.frx":8932
      colorPressed    =   "frmCapture.frx":895C
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   540
      MaxLength       =   4
      TabIndex        =   66
      Text            =   "0"
      Top             =   7620
      Width           =   945
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   65
      Text            =   "0.00"
      Top             =   7125
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   64
      Text            =   "0.00"
      Top             =   7620
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   "0.00"
      Top             =   8460
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "0.00"
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "0.00"
      Top             =   6075
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "0.00"
      Top             =   5550
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "0.00"
      Top             =   5025
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "0.00"
      Top             =   4500
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "0.00"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "0.00"
      Top             =   3075
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "0.00"
      Top             =   2550
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "0.00"
      Top             =   2025
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "0.00"
      Top             =   1470
      Width           =   1455
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   2490
      MaxLength       =   12
      TabIndex        =   52
      Text            =   "0.00"
      Top             =   10020
      Width           =   1845
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   2490
      MaxLength       =   12
      TabIndex        =   51
      Text            =   "0.00"
      Top             =   9495
      Width           =   1845
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   2490
      MaxLength       =   12
      TabIndex        =   50
      Text            =   "0.00"
      Top             =   8970
      Width           =   1845
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   540
      MaxLength       =   4
      TabIndex        =   49
      Text            =   "0"
      Top             =   7125
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   540
      MaxLength       =   4
      TabIndex        =   48
      Text            =   "0"
      Top             =   6600
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   540
      MaxLength       =   4
      TabIndex        =   47
      Text            =   "0"
      Top             =   6075
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   540
      MaxLength       =   4
      TabIndex        =   46
      Text            =   "0"
      Top             =   5550
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   540
      MaxLength       =   4
      TabIndex        =   45
      Text            =   "0"
      Top             =   5025
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   540
      MaxLength       =   4
      TabIndex        =   44
      Text            =   "0"
      Top             =   4500
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   540
      MaxLength       =   4
      TabIndex        =   43
      Text            =   "0"
      Top             =   3600
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   540
      MaxLength       =   4
      TabIndex        =   42
      Text            =   "0"
      Top             =   3075
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   540
      MaxLength       =   4
      TabIndex        =   41
      Text            =   "0"
      Top             =   2550
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   540
      MaxLength       =   4
      TabIndex        =   40
      Text            =   "0"
      Top             =   2025
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   540
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "0"
      Top             =   1470
      Width           =   945
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1425
      Index           =   9
      Left            =   4650
      TabIndex        =   1
      Top             =   7905
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2514
      _StockProps     =   66
      Caption         =   "CL"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   24
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
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCapture.frx":8986
      textLT          =   "frmCapture.frx":89EA
      textCT          =   "frmCapture.frx":8A02
      textRT          =   "frmCapture.frx":8A1A
      textLM          =   "frmCapture.frx":8A32
      textRM          =   "frmCapture.frx":8A4A
      textLB          =   "frmCapture.frx":8A62
      textCB          =   "frmCapture.frx":8A7A
      textRB          =   "frmCapture.frx":8A92
      colorBack       =   "frmCapture.frx":8AAA
      colorIntern     =   "frmCapture.frx":8AD4
      colorMO         =   "frmCapture.frx":8AFE
      colorFocus      =   "frmCapture.frx":8B28
      colorDisabled   =   "frmCapture.frx":8B52
      colorPressed    =   "frmCapture.frx":8B7C
      Orientation     =   8
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1395
      Index           =   6
      Left            =   4650
      TabIndex        =   4
      Top             =   6510
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2461
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
      textCaption     =   "frmCapture.frx":8BA6
      textLT          =   "frmCapture.frx":8C08
      textCT          =   "frmCapture.frx":8C20
      textRT          =   "frmCapture.frx":8C38
      textLM          =   "frmCapture.frx":8C50
      textRM          =   "frmCapture.frx":8C68
      textLB          =   "frmCapture.frx":8C80
      textCB          =   "frmCapture.frx":8C98
      textRB          =   "frmCapture.frx":8CB0
      colorBack       =   "frmCapture.frx":8CC8
      colorIntern     =   "frmCapture.frx":8CF2
      colorMO         =   "frmCapture.frx":8D1C
      colorFocus      =   "frmCapture.frx":8D46
      colorDisabled   =   "frmCapture.frx":8D70
      colorPressed    =   "frmCapture.frx":8D9A
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1335
      Index           =   3
      Left            =   4650
      TabIndex        =   7
      Top             =   5175
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2355
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
      textCaption     =   "frmCapture.frx":8DC4
      textLT          =   "frmCapture.frx":8E26
      textCT          =   "frmCapture.frx":8E3E
      textRT          =   "frmCapture.frx":8E56
      textLM          =   "frmCapture.frx":8E6E
      textRM          =   "frmCapture.frx":8E86
      textLB          =   "frmCapture.frx":8E9E
      textCB          =   "frmCapture.frx":8EB6
      textRB          =   "frmCapture.frx":8ECE
      colorBack       =   "frmCapture.frx":8EE6
      colorIntern     =   "frmCapture.frx":8F10
      colorMO         =   "frmCapture.frx":8F3A
      colorFocus      =   "frmCapture.frx":8F64
      colorDisabled   =   "frmCapture.frx":8F8E
      colorPressed    =   "frmCapture.frx":8FB8
      Orientation     =   5
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1425
      Index           =   10
      Left            =   6165
      TabIndex        =   2
      Top             =   7905
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
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
      textCaption     =   "frmCapture.frx":8FE2
      textLT          =   "frmCapture.frx":9044
      textCT          =   "frmCapture.frx":905C
      textRT          =   "frmCapture.frx":9074
      textLM          =   "frmCapture.frx":908C
      textRM          =   "frmCapture.frx":90A4
      textLB          =   "frmCapture.frx":90BC
      textCB          =   "frmCapture.frx":90D4
      textRB          =   "frmCapture.frx":90EC
      colorBack       =   "frmCapture.frx":9104
      colorIntern     =   "frmCapture.frx":912E
      colorMO         =   "frmCapture.frx":9158
      colorFocus      =   "frmCapture.frx":9182
      colorDisabled   =   "frmCapture.frx":91AC
      colorPressed    =   "frmCapture.frx":91D6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1395
      Index           =   7
      Left            =   6165
      TabIndex        =   5
      Top             =   6510
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2461
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
      textCaption     =   "frmCapture.frx":9200
      textLT          =   "frmCapture.frx":9262
      textCT          =   "frmCapture.frx":927A
      textRT          =   "frmCapture.frx":9292
      textLM          =   "frmCapture.frx":92AA
      textRM          =   "frmCapture.frx":92C2
      textLB          =   "frmCapture.frx":92DA
      textCB          =   "frmCapture.frx":92F2
      textRB          =   "frmCapture.frx":930A
      colorBack       =   "frmCapture.frx":9322
      colorIntern     =   "frmCapture.frx":934C
      colorMO         =   "frmCapture.frx":9376
      colorFocus      =   "frmCapture.frx":93A0
      colorDisabled   =   "frmCapture.frx":93CA
      colorPressed    =   "frmCapture.frx":93F4
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1335
      Index           =   4
      Left            =   6165
      TabIndex        =   8
      Top             =   5175
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2355
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
      textCaption     =   "frmCapture.frx":941E
      textLT          =   "frmCapture.frx":9480
      textCT          =   "frmCapture.frx":9498
      textRT          =   "frmCapture.frx":94B0
      textLM          =   "frmCapture.frx":94C8
      textRM          =   "frmCapture.frx":94E0
      textLB          =   "frmCapture.frx":94F8
      textCB          =   "frmCapture.frx":9510
      textRB          =   "frmCapture.frx":9528
      colorBack       =   "frmCapture.frx":9540
      colorIntern     =   "frmCapture.frx":956A
      colorMO         =   "frmCapture.frx":9594
      colorFocus      =   "frmCapture.frx":95BE
      colorDisabled   =   "frmCapture.frx":95E8
      colorPressed    =   "frmCapture.frx":9612
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1425
      Index           =   11
      Left            =   7680
      TabIndex        =   3
      Top             =   7905
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2514
      _StockProps     =   66
      Caption         =   "OK"
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
      CornerFactor    =   15
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCapture.frx":963C
      textLT          =   "frmCapture.frx":96A0
      textCT          =   "frmCapture.frx":96B8
      textRT          =   "frmCapture.frx":96D0
      textLM          =   "frmCapture.frx":96E8
      textRM          =   "frmCapture.frx":9700
      textLB          =   "frmCapture.frx":9718
      textCB          =   "frmCapture.frx":9730
      textRB          =   "frmCapture.frx":9748
      colorBack       =   "frmCapture.frx":9760
      colorIntern     =   "frmCapture.frx":978A
      colorMO         =   "frmCapture.frx":97B4
      colorFocus      =   "frmCapture.frx":97DE
      colorDisabled   =   "frmCapture.frx":9808
      colorPressed    =   "frmCapture.frx":9832
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1395
      Index           =   8
      Left            =   7680
      TabIndex        =   6
      Top             =   6510
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2461
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
      textCaption     =   "frmCapture.frx":985C
      textLT          =   "frmCapture.frx":98BE
      textCT          =   "frmCapture.frx":98D6
      textRT          =   "frmCapture.frx":98EE
      textLM          =   "frmCapture.frx":9906
      textRM          =   "frmCapture.frx":991E
      textLB          =   "frmCapture.frx":9936
      textCB          =   "frmCapture.frx":994E
      textRB          =   "frmCapture.frx":9966
      colorBack       =   "frmCapture.frx":997E
      colorIntern     =   "frmCapture.frx":99A8
      colorMO         =   "frmCapture.frx":99D2
      colorFocus      =   "frmCapture.frx":99FC
      colorDisabled   =   "frmCapture.frx":9A26
      colorPressed    =   "frmCapture.frx":9A50
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1335
      Index           =   5
      Left            =   7680
      TabIndex        =   9
      Top             =   5175
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2355
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
      textCaption     =   "frmCapture.frx":9A7A
      textLT          =   "frmCapture.frx":9ADC
      textCT          =   "frmCapture.frx":9AF4
      textRT          =   "frmCapture.frx":9B0C
      textLM          =   "frmCapture.frx":9B24
      textRM          =   "frmCapture.frx":9B3C
      textLB          =   "frmCapture.frx":9B54
      textCB          =   "frmCapture.frx":9B6C
      textRB          =   "frmCapture.frx":9B84
      colorBack       =   "frmCapture.frx":9B9C
      colorIntern     =   "frmCapture.frx":9BC6
      colorMO         =   "frmCapture.frx":9BF0
      colorFocus      =   "frmCapture.frx":9C1A
      colorDisabled   =   "frmCapture.frx":9C44
      colorPressed    =   "frmCapture.frx":9C6E
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   1
      Left            =   8250
      TabIndex        =   18
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   2163158
      Caption         =   "X"
      CaptionOffsetY  =   2
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdKey 
      Height          =   870
      Index           =   0
      Left            =   6450
      TabIndex        =   19
      Top             =   240
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1535
      Appearance      =   3
      BackColor       =   10736617
      Caption         =   "Accept"
      CaptionOffsetY  =   2
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1395
      Index           =   1
      Left            =   6165
      TabIndex        =   27
      Top             =   3780
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2461
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
      textCaption     =   "frmCapture.frx":9C98
      textLT          =   "frmCapture.frx":9CFA
      textCT          =   "frmCapture.frx":9D12
      textRT          =   "frmCapture.frx":9D2A
      textLM          =   "frmCapture.frx":9D42
      textRM          =   "frmCapture.frx":9D5A
      textLB          =   "frmCapture.frx":9D72
      textCB          =   "frmCapture.frx":9D8A
      textRB          =   "frmCapture.frx":9DA2
      colorBack       =   "frmCapture.frx":9DBA
      colorIntern     =   "frmCapture.frx":9DE4
      colorMO         =   "frmCapture.frx":9E0E
      colorFocus      =   "frmCapture.frx":9E38
      colorDisabled   =   "frmCapture.frx":9E62
      colorPressed    =   "frmCapture.frx":9E8C
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1395
      Index           =   2
      Left            =   7680
      TabIndex        =   28
      Top             =   3780
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2461
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
      CornerFactor    =   15
      BackColorContainer=   12632256
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCapture.frx":9EB6
      textLT          =   "frmCapture.frx":9F18
      textCT          =   "frmCapture.frx":9F30
      textRT          =   "frmCapture.frx":9F48
      textLM          =   "frmCapture.frx":9F60
      textRM          =   "frmCapture.frx":9F78
      textLB          =   "frmCapture.frx":9F90
      textCB          =   "frmCapture.frx":9FA8
      textRB          =   "frmCapture.frx":9FC0
      colorBack       =   "frmCapture.frx":9FD8
      colorIntern     =   "frmCapture.frx":A002
      colorMO         =   "frmCapture.frx":A02C
      colorFocus      =   "frmCapture.frx":A056
      colorDisabled   =   "frmCapture.frx":A080
      colorPressed    =   "frmCapture.frx":A0AA
      Orientation     =   6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1245
      Index           =   12
      Left            =   4650
      TabIndex        =   37
      Top             =   9330
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2196
      _StockProps     =   66
      Caption         =   "«"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   36
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
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCapture.frx":A0D4
      textLT          =   "frmCapture.frx":A136
      textCT          =   "frmCapture.frx":A14E
      textRT          =   "frmCapture.frx":A166
      textLM          =   "frmCapture.frx":A17E
      textRM          =   "frmCapture.frx":A196
      textLB          =   "frmCapture.frx":A1AE
      textCB          =   "frmCapture.frx":A1C6
      textRB          =   "frmCapture.frx":A1DE
      colorBack       =   "frmCapture.frx":A1F6
      colorIntern     =   "frmCapture.frx":A220
      colorMO         =   "frmCapture.frx":A24A
      colorFocus      =   "frmCapture.frx":A274
      colorDisabled   =   "frmCapture.frx":A29E
      colorPressed    =   "frmCapture.frx":A2C8
      Orientation     =   8
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1245
      Index           =   13
      Left            =   6165
      TabIndex        =   38
      Top             =   9330
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2196
      _StockProps     =   66
      Caption         =   "."
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
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
      textCaption     =   "frmCapture.frx":A2F2
      textLT          =   "frmCapture.frx":A354
      textCT          =   "frmCapture.frx":A36C
      textRT          =   "frmCapture.frx":A384
      textLM          =   "frmCapture.frx":A39C
      textRM          =   "frmCapture.frx":A3B4
      textLB          =   "frmCapture.frx":A3CC
      textCB          =   "frmCapture.frx":A3E4
      textRB          =   "frmCapture.frx":A3FC
      colorBack       =   "frmCapture.frx":A414
      colorIntern     =   "frmCapture.frx":A43E
      colorMO         =   "frmCapture.frx":A468
      colorFocus      =   "frmCapture.frx":A492
      colorDisabled   =   "frmCapture.frx":A4BC
      colorPressed    =   "frmCapture.frx":A4E6
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin BTNENHLib4.BtnEnh cmdInput 
      Height          =   1245
      Index           =   14
      Left            =   7680
      TabIndex        =   39
      Top             =   9330
      Width           =   1545
      _Version        =   524298
      _ExtentX        =   2725
      _ExtentY        =   2196
      _StockProps     =   66
      Caption         =   "»"
      BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   36
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
      SpecialEffect   =   1
      LogPixels       =   96
      SpecialEffectFactor=   4
      TextureBevelFactor=   4
      UserData        =   0.1
      textCaption     =   "frmCapture.frx":A510
      textLT          =   "frmCapture.frx":A572
      textCT          =   "frmCapture.frx":A58A
      textRT          =   "frmCapture.frx":A5A2
      textLM          =   "frmCapture.frx":A5BA
      textRM          =   "frmCapture.frx":A5D2
      textLB          =   "frmCapture.frx":A5EA
      textCB          =   "frmCapture.frx":A602
      textRB          =   "frmCapture.frx":A61A
      colorBack       =   "frmCapture.frx":A632
      colorIntern     =   "frmCapture.frx":A65C
      colorMO         =   "frmCapture.frx":A686
      colorFocus      =   "frmCapture.frx":A6B0
      colorDisabled   =   "frmCapture.frx":A6DA
      colorPressed    =   "frmCapture.frx":A704
      Orientation     =   7
      HollowFrame     =   -1  'True
      LightDirection  =   7
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   26
      Left            =   2820
      Top             =   1410
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   25
      Left            =   2820
      Top             =   1935
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   24
      Left            =   2820
      Top             =   2460
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   23
      Left            =   2820
      Top             =   2985
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   22
      Left            =   2820
      Top             =   3510
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   480
      Index           =   27
      Left            =   2280
      Top             =   8400
      Width           =   2130
      BorderColor     =   8421504
      Size            =   "3757;855"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   510
      Index           =   21
      Left            =   2820
      Top             =   4380
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;900"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   510
      Index           =   20
      Left            =   2820
      Top             =   4905
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;900"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   510
      Index           =   19
      Left            =   2820
      Top             =   5430
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;900"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   510
      Index           =   18
      Left            =   2820
      Top             =   5955
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;900"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   510
      Index           =   17
      Left            =   2820
      Top             =   6480
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;900"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   510
      Index           =   16
      Left            =   2820
      Top             =   7005
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;900"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   510
      Index           =   15
      Left            =   2820
      Top             =   7530
      Width           =   1575
      BorderColor     =   8421504
      Size            =   "2778;900"
      VariousPropertyBits=   19
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   8
      Left            =   1440
      TabIndex        =   67
      Top             =   7125
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X 10c ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   14
      Left            =   2280
      Top             =   9990
      Width           =   2130
      BorderColor     =   8421504
      Size            =   "3766;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   13
      Left            =   2280
      Top             =   9465
      Width           =   2130
      BorderColor     =   8421504
      Size            =   "3766;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   12
      Left            =   2280
      Top             =   8925
      Width           =   2130
      BorderColor     =   8421504
      Size            =   "3766;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   515
      Index           =   11
      Left            =   420
      Top             =   7530
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;908"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   515
      Index           =   10
      Left            =   420
      Top             =   7005
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;908"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   515
      Index           =   9
      Left            =   420
      Top             =   6480
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;908"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   515
      Index           =   8
      Left            =   420
      Top             =   5955
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;908"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   515
      Index           =   7
      Left            =   420
      Top             =   5430
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;908"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   515
      Index           =   6
      Left            =   420
      Top             =   4905
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;908"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   515
      Index           =   5
      Left            =   420
      Top             =   4380
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;908"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   4
      Left            =   420
      Top             =   3510
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   3
      Left            =   420
      Top             =   2985
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   480
      Index           =   2
      Left            =   420
      Top             =   2460
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;847"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   1
      Left            =   420
      Top             =   1935
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image6 
      Height          =   495
      Index           =   0
      Left            =   420
      Top             =   1410
      Width           =   1095
      BorderColor     =   8421504
      Size            =   "1931;873"
      VariousPropertyBits=   19
   End
   Begin MSForms.Image Image5 
      Height          =   615
      Left            =   1200
      Top             =   8310
      Width           =   2865
      BorderStyle     =   0
      Size            =   "5054;1085"
      VariousPropertyBits=   19
   End
   Begin MSForms.Label lblHeading 
      Height          =   585
      Left            =   570
      TabIndex        =   36
      Top             =   510
      Width           =   6255
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Heading"
      Size            =   "11033;1032"
      FontName        =   "Arial Narrow"
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblTot 
      Height          =   525
      Left            =   4980
      TabIndex        =   35
      Top             =   2910
      Width           =   1755
      VariousPropertyBits=   8388627
      Caption         =   "Variance:"
      Size            =   "3096;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblRep 
      Height          =   525
      Left            =   4980
      TabIndex        =   34
      Top             =   2265
      Width           =   1755
      VariousPropertyBits=   8388627
      Caption         =   "Reported Total:"
      Size            =   "3096;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   585
      Index           =   14
      Left            =   4980
      TabIndex        =   33
      Top             =   1620
      Width           =   1755
      VariousPropertyBits=   8388627
      Caption         =   "Count Total:"
      Size            =   "3096;1032"
      FontName        =   "Arial Narrow"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   13
      Left            =   630
      TabIndex        =   32
      Top             =   10035
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "Charge Total:"
      Size            =   "2725;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   12
      Left            =   630
      TabIndex        =   31
      Top             =   9510
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "Voucher Total:"
      Size            =   "2725;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   11
      Left            =   630
      TabIndex        =   30
      Top             =   8985
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "Card Total:"
      Size            =   "2725;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   10
      Left            =   630
      TabIndex        =   29
      Top             =   8460
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "Cash Total:"
      Size            =   "2725;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   9
      Left            =   1440
      TabIndex        =   25
      Top             =   7650
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X 5c ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   7
      Left            =   1440
      TabIndex        =   24
      Top             =   6600
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X 20c ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   6
      Left            =   1440
      TabIndex        =   23
      Top             =   6075
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X 50c ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   5
      Left            =   1440
      TabIndex        =   22
      Top             =   5550
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X R1-00 ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   525
      Index           =   1
      Left            =   1440
      TabIndex        =   21
      Top             =   5025
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X R2-00 ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   4
      Left            =   1440
      TabIndex        =   20
      Top             =   4500
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X R5-00 ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   0
      Left            =   1440
      TabIndex        =   17
      Top             =   1530
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X R200-00 ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   525
      Index           =   0
      Left            =   1440
      TabIndex        =   16
      Top             =   2055
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X R100-00 ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   1
      Left            =   1440
      TabIndex        =   15
      Top             =   2580
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X R50-00 ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   2
      Left            =   1440
      TabIndex        =   14
      Top             =   3105
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X R20-00 ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Index           =   3
      Left            =   1440
      TabIndex        =   13
      Top             =   3630
      Width           =   1485
      VariousPropertyBits=   8388627
      Caption         =   "X R10-00 ="
      Size            =   "2619;926"
      FontName        =   "Arial Narrow"
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblCount 
      Height          =   645
      Index           =   2
      Left            =   6960
      TabIndex        =   12
      Top             =   2910
      Width           =   2055
      BackColor       =   15461355
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "3625;1138"
      FontName        =   "Arial Narrow"
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCount 
      Height          =   645
      Index           =   1
      Left            =   6960
      TabIndex        =   11
      Top             =   2250
      Width           =   2055
      BackColor       =   15461355
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "3625;1138"
      FontName        =   "Arial Narrow"
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label lblCount 
      Height          =   645
      Index           =   0
      Left            =   6960
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
      BackColor       =   15461355
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "3625;1138"
      FontName        =   "Arial Narrow"
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Image Image3 
      Height          =   615
      Left            =   6810
      Top             =   2790
      Width           =   2325
      BackColor       =   11196396
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "4101;1085"
   End
   Begin MSForms.Image pic 
      Height          =   615
      Left            =   6810
      Top             =   2115
      Width           =   2325
      BackColor       =   14737632
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "4101;1085"
   End
   Begin MSForms.Image Image1 
      Height          =   615
      Left            =   6810
      Top             =   1440
      Width           =   2325
      BackColor       =   11196396
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "4101;1085"
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
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdInput_Click(Index As Integer)
    If cmdKey(0).Enabled = False Then
        Exit Sub
    End If
    Select Case cmdInput(Index).Caption
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
            Select Case frmCapture.Tag
                Case 0 To 11
                    If txtQty(frmCapture.Tag).Text = "0" Then txtQty(frmCapture.Tag).Text = ""
                    If txtQty(frmCapture.Tag).SelLength = Len(txtQty(frmCapture.Tag).Text) Then txtQty(frmCapture.Tag).Text = ""
                Case 16
                    If txtFloat.SelLength = Len(txtFloat.Text) Then txtFloat.Text = ""
                    If txtFloat.Text = "0" Then txtFloat.Text = ""
                    If Left(txtFloat.Text, 4) = "0.00" Then txtFloat.Text = Replace(txtFloat.Text, "0.00", "")
                Case Else
                    If txtQty(frmCapture.Tag).SelLength = Len(txtQty(frmCapture.Tag).Text) Then txtQty(frmCapture.Tag).Text = ""
                    If txtQty(frmCapture.Tag).Text = "0" Then txtQty(frmCapture.Tag).Text = ""
                    If Left(txtQty(frmCapture.Tag).Text, 4) = "0.00" Then txtQty(frmCapture.Tag).Text = Replace(txtQty(frmCapture.Tag).Text, "0.00", "")
            End Select
            If frmCapture.Tag = 16 Then
                txtFloat.Tag = "1"
                txtFloat.SetFocus
                txtFloat.Tag = ""
                SendKeys "{END}"
                SendKeys cmdInput(Index).Caption
            Else
                txtQty(frmCapture.Tag).Tag = "1"
                txtQty(frmCapture.Tag).SetFocus
                txtQty(frmCapture.Tag).Tag = ""
                SendKeys "{END}"
                SendKeys cmdInput(Index).Caption
            DoEvents
            End If
        Case "CL"
            If frmCapture.Tag = 16 Then
                txtFloat.Tag = ""
                txtFloat.SetFocus
            Else
                txtQty(frmCapture.Tag).Tag = ""
                txtQty(frmCapture.Tag).SetFocus
            End If
            Select Case Val(frmCapture.Tag)
                Case 0 To 11
                    txtQty(frmCapture.Tag).Text = "0"
                Case 16
                    txtFloat.Text = "0.00"
                Case Else
                    txtQty(frmCapture.Tag).Text = "0.00"
            End Select
            Calculate
        Case "OK"
            Select Case Val(frmCapture.Tag)
                Case 12: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 13: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 14: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 16: txtFloat.Text = Format(txtFloat.Text, "0.00")
            End Select
            Calculate
            If Val(frmCapture.Tag) = 14 Then frmCapture.Tag = -1
            If Val(frmCapture.Tag) = -1 And lblTot.Caption = "Total:" Then
                txtFloat.SetFocus
                Exit Sub
            End If
            If Val(frmCapture.Tag) = 16 And lblTot.Caption = "Total:" Then
                frmCapture.Tag = -1
            End If
            txtQty(Val(frmCapture.Tag) + 1).SetFocus
        Case "«"
            Select Case Val(frmCapture.Tag)
                Case 12: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 13: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 14: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 16: txtFloat.Text = Format(txtFloat.Text, "0.00")
            End Select
            Calculate
            If Val(frmCapture.Tag) = 0 Then frmCapture.Tag = 15
            If Val(frmCapture.Tag) = 15 And lblTot.Caption = "Total:" Then
                txtFloat.SetFocus
                Exit Sub
            End If
            If Val(frmCapture.Tag) = 16 And lblTot.Caption = "Total:" Then
                frmCapture.Tag = 15
            End If
            txtQty(Val(frmCapture.Tag) - 1).SetFocus
        Case "»"
            Select Case Val(frmCapture.Tag)
                Case 12: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 13: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 14: txtQty(Val(frmCapture.Tag)).Text = Format(txtQty(Val(frmCapture.Tag)).Text, "0.00")
                Case 16: txtFloat.Text = Format(txtFloat.Text, "0.00")
            End Select
            Calculate
            If Val(frmCapture.Tag) = 14 Then frmCapture.Tag = -1
            If Val(frmCapture.Tag) = -1 And lblTot.Caption = "Total:" Then
                txtFloat.SetFocus
                Exit Sub
            End If
            If Val(frmCapture.Tag) = 16 And lblTot.Caption = "Total:" Then
                frmCapture.Tag = -1
            End If
            txtQty(Val(frmCapture.Tag) + 1).SetFocus
    End Select
End Sub
Private Sub Calculate()
    If Val(frmCapture.Tag) = -1 Then Exit Sub
    If Val(frmCapture.Tag) < 12 Then
        txtQty(Val(frmCapture.Tag)).Text = Val(txtQty(Val(frmCapture.Tag)).Text)
    End If
    Select Case Val(frmCapture.Tag)
        Case 0: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 200, "0.00")
        Case 1: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 100, "0.00")
        Case 2: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 50, "0.00")
        Case 3: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 20, "0.00")
        Case 4: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 10, "0.00")
        Case 5: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 5, "0.00")
        Case 6: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 2, "0.00")
        Case 7: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 1, "0.00")
        Case 8: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 0.5, "0.00")
        Case 9: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 0.2, "0.00")
        Case 10: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 0.1, "0.00")
        Case 11: txtValue(Val(frmCapture.Tag)).Text = Format(Val(txtQty(Val(frmCapture.Tag))) * 0.05, "0.00")
    End Select
    If Val(frmCapture.Tag) < 12 Then
        If txtQty(Val(frmCapture.Tag)).Text = "" Then
            txtQty(Val(frmCapture.Tag)).Text = "0"
        End If
    End If
    txtValue(13).Text = "0.00"
    For i = 0 To 11
        txtValue(13).Text = Format(Val(txtValue(13).Text) + Val(txtValue(i).Text), "0.00")
    Next i
    
'     If rs.State = 1 Then
'    lblCount(0).Caption = Format(rs.Fields("Counted") & "", "0.00")
'    Else
    lblCount(0).Caption = Format(Val(txtValue(13).Text) + Val(txtQty(12).Text) + Val(txtQty(13).Text) + Val(txtQty(14).Text), "0.00")
'    End If
    If lblTot.Caption = "Total:" Then
    'lblCount(0).Caption = Format(rs.Fields("Counted") & "", "0.00")
        
        lblCount(2).Caption = Format(Val(lblCount(0).Caption) - Val(txtFloat.Text), "0.00")
    Else
        'lblCount(0).Caption = Format(rs.Fields("Counted") & "", "0.00")
        
        lblCount(2).Caption = Format(Val(lblCount(0).Caption) - Val(lblCount(1).Caption), "0.00")
    End If
End Sub
Private Sub cmdKey_Click(Index As Integer)
    Select Case Index
        Case 1
            Unload Me
        Case Else
            
            Screen.MousePointer = 11
            If lblTot.Caption = "Total:" Then
               
                
                
                ActiveUpdateServer "Update Counters set " & _
                "[200]=" & Val(txtQty(0).Text) & ",[100]=" & Val(txtQty(1).Text) & _
                ",[50]=" & Val(txtQty(2).Text) & ",[20]=" & Val(txtQty(3).Text) & ",[10]=" & _
                Val(txtQty(4).Text) & ",[5]=" & Val(txtQty(5).Text) & ",[2]=" & Val(txtQty(6).Text) & _
                ",[1]=" & Val(txtQty(7).Text) & ",[50c]=" & Val(txtQty(8).Text) & ",[20c]=" & _
                Val(txtQty(9).Text) & ",[10c]=" & Val(txtQty(10).Text) & ",[5c]=" & Val(txtQty(11).Text) & _
                ",[CardC]=" & Val(txtQty(12).Text) & ",[ChequeC]=" & Val(txtQty(13).Text) & ",[ChargeC]=" & _
                Val(txtQty(14).Text) & ",Counted = " & lblCount(2).Caption & ",Date_Time=Getdate() " & _
                " where Cashup_No= " & frmCapture.lblHeading.Tag
            Else
                 Debug.Print lblCount(2).Caption
                 Debug.Print lblCount(0).Caption
                 
                ActiveUpdateServer "Update Counters set " & _
                "[200]=" & Val(txtQty(0).Text) & ",[100]=" & Val(txtQty(1).Text) & _
                ",[50]=" & Val(txtQty(2).Text) & ",[20]=" & Val(txtQty(3).Text) & _
                ",[10]=" & Val(txtQty(4).Text) & ",[5]=" & Val(txtQty(5).Text) & _
                ",[2]=" & Val(txtQty(6).Text) & ",[1]=" & Val(txtQty(7).Text) & _
                ",[50c]=" & Val(txtQty(8).Text) & ",[20c]=" & Val(txtQty(9).Text) & _
                ",[10c]=" & Val(txtQty(10).Text) & ",[5c]=" & Val(txtQty(11).Text) & _
                ",[CardC]=" & Val(txtQty(12).Text) & ",[ChequeC]=" & Val(txtQty(13).Text) & _
                ",[ChargeC]=" & Val(txtQty(14).Text) & ",Counted = " & lblCount(0).Caption & " " & _
                " where Cashup_No= " & frmTillReport.lblCashupInfo.Tag
                frmTillReport.lblCash(48).Caption = lblCount(0).Caption
                frmTillReport.lblCash(49) = Format(Val(frmTillReport.lblCash(48)) - Val(frmTillReport.lblCash(47)), "0.00")
            End If
            PrintCountSlip1
            Seeiffinalizing
            frmTillReport.Tag = "Not Now"
            Unload Me
            Screen.MousePointer = 0
    End Select
End Sub
Private Sub Seeiffinalizing()
If Senttofinalize = True Then
 
        









ActiveReadServer "Select count(Table_No) as TableCount" & _
            " from Table_Listing_View where User_No = " & UserRecord.User_Number
            If rs.Fields("TableCount") > 0 Then
                Timer2.Enabled = True
                Select Case rs.Fields("TableCount")
                    Case 1
                        cmdErr.Caption = "This User still has an Open Table."
                    Case Else
                        cmdErr.Caption = "This User still has " & rs.Fields("TableCount") & " Open Tables."
                End Select
                cmdErr.Visible = True
                rs.Close
                Screen.MousePointer = 1
                Exit Sub
            End If
            rs.Close
            
            ActiveReadServer "Select count(Tab_No) as TabCount from Tab_Listing_View where User_No = " & UserRecord.User_Number
            If rs.Fields("TabCount") > 0 Then
                Timer2.Enabled = True
                Select Case rs.Fields("TabCount")
                    Case 1
                        cmdErr.Caption = "This User still has an Open Tab."
                    Case Else
                        cmdErr.Caption = "This User still has " & rs.Fields("TabCount") & " Open Tabs."
                End Select
                cmdErr.Visible = True
                rs.Close
                Screen.MousePointer = 1
                Exit Sub
            End If
            rs.Close
            
            ActiveReadServer "Select count(Table_No) as TableCount from Table_Listing_View where User_No <> " & UserRecord.User_Number & _
            " and Previous_Owner = " & UserRecord.User_Number
            If rs.Fields("TableCount") > 0 Then
                Timer2.Enabled = True
                Select Case rs.Fields("TableCount")
                    Case 1
                        cmdErr.Caption = "A Transfer by this User has nor been accepted."
                    Case Else
                        cmdErr.Caption = rs.Fields("TableCount") & " Transfers by this User has nor been accepted."
                End Select
                cmdErr.Visible = True
                rs.Close
                Screen.MousePointer = 1
                Exit Sub
            End If
            rs.Close
            FinalizingCashup
            Screen.MousePointer = 1



End If
'ActiveReadServer "Select Date_Time from Counters where Cashup_No = " & frmTillReport.lblCashupInfo.Tag
Senttofinalize = False
End Sub

Private Sub FinalizingCashup()
Dim opentrans As String, closetrans As String
ActiveReadServer1 "Select (Select min(Invoice_No) from Sales_Journal where Cashup_No=" & frmCapture.lblHeading.Tag & " and User_No=" & UserRecord.User_Number & ") as Open_Trans," & _
        "(Select max(Invoice_No) from Sales_Journal where Cashup_No=" & frmCapture.lblHeading.Tag & " and User_No=" & UserRecord.User_Number & ") as Close_Trans"
        
    opentrans = Val(rs1.Fields("Open_Trans") & "")
    closetrans = Val(rs1.Fields("Close_Trans") & "")
    rs1.Close
    
    ActiveReadServer1 "Select (Select min(Date_time) from Sales_Journal where Cashup_No=" & frmCapture.lblHeading.Tag & " and User_No=" & UserRecord.User_Number & ") as Start_time," & _
        "(Select max(Date_Time) from Sales_Journal where Cashup_No=" & frmCapture.lblHeading.Tag & " and User_No=" & UserRecord.User_Number & ") as End_Time"
     
    
    startdatetime = rs1.Fields("Start_Time")
    enddatetime = rs1.Fields("End_Time")
rs1.Close
ActiveUpdateServer "Update Users set Drawer_No = 0, Clocked_In = 0 where User_No = " & UserRecord.User_Number
            
           If rs.State = 1 Then rs.Close
            ActiveReadServer "Select Date_Time from Counters where Cashup_No= " & frmCapture.lblHeading.Tag
            If rs.Fields("Date_Time") & "" = "" Then
                rs.Close
                ActiveUpdateServer "Update Counters set Date_Time=Getdate(),Finalized=1,Workstation_No=" & Workstation_No & ",Open_Trans_No='" & opentrans & "',Close_Trans_No='" & closetrans & "',Shift_Start = '" & startdatetime & "' where Cashup_No= " & frmCapture.lblHeading.Tag
            Else
                rs.Close
                ActiveUpdateServer "Update Counters set Finalized=1,Workstation_No=" & Workstation_No & ",Open_Trans_No='" & opentrans & "',Close_Trans_No='" & closetrans & "',Shift_Start = '" & startdatetime & "' where Cashup_No= " & frmCapture.lblHeading.Tag
            End If
End Sub




Private Sub PrintCountSlip1()
    On Error GoTo trap
    Dim x As Printer
    PrintErr = 0
    Slip_Port = ""
    filenum = FreeFile
    Close #filenum
    If Slip_PrinterPort = 0 Then ' Kotie 17-03-2013
        If InStr(Trim(Slip_Printer), "\\") = 0 Then
            If Slip_Port = "" Then
                Open "\\" & Comp_Name & "\" & Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        Else
            If Slip_Port = "" Then
                Open Slip_Printer For Output As filenum
            Else
                Open Slip_Port For Output As filenum
            End If
        End If
        If Slip_Port <> "" Then
            If UCase(Left(Slip_Port, 2)) = "NE" Then
                Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
            Else
                If Slip_Port = "FILE:" Then
                    Open "C:\" & x.DeviceName & ".txt" For Output As filenum
                Else
                    Open Slip_Port For Output As filenum
                End If
            End If
        End If
    Else
        Open "Com" & Trim(Slip_PrinterPort) & ":" For Output As filenum
    End If
    Print #filenum, Chr(27) & Chr(64);
    Print #filenum, Chr(27) & Chr(69) & Chr(1);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, UCase(Branch_Name)
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(33) & Chr(16);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "CASH UP COUNT"
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(49);
    Print #filenum, String(40, "=")
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(97) & Chr(49);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, "CASHUP NO: " & frmCapture.lblHeading.Tag
    Print #filenum, UCase(Trim(UserRecord.User_Number) & " - " & Trim(UserRecord.Name))
    Print #filenum, UCase(Format(Date, "dd MMMM yyyy DDD") & " " & Format(Time, "HH:MM:SS"))
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(69) & Chr(48);
    Print #filenum, Chr(27) & Chr(77) & Chr(48);
    Print #filenum, Chr(27) & Chr(33) & Chr(0);
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(27) & Chr(97) & Chr(50);
    Print #filenum, Chr(27) & Chr(51) & Chr(18);
    Print #filenum, Chr(27) & Chr(69) & Chr(49);
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CASH TOTAL:  " & Chr(179) & String(14 - Len(txtValue(13)), Chr(32)) & txtValue(13) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CARD TOTAL:  " & Chr(179) & String(14 - Len(txtQty(12)), Chr(32)) & txtQty(12) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "VOUCHER TOTAL:  " & Chr(179) & String(14 - Len(txtQty(13)), Chr(32)) & txtQty(13) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "CHARGE TOTAL:  " & Chr(179) & String(14 - Len(txtQty(14)), Chr(32)) & txtQty(14) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, String(33, "-")
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "COUNT TOTAL:  " & Chr(179) & String(14 - Len(lblCount(0)), Chr(32)) & lblCount(0) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "      -FLOAT:  " & Chr(179) & String(14 - Len(txtFloat.Text), Chr(32)) & txtFloat.Text & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, Chr(218) & String(15, Chr(196)) & Chr(191)
    Print #filenum, "" & Chr(179) & String(15, Chr(32)) & Chr(179)
    Print #filenum, "     TOTAL:  " & Chr(179) & String(14 - Len(lblCount(2)), Chr(32)) & lblCount(2) & " " & Chr(179)
    Print #filenum, Chr(192) & String(15, Chr(196)) & Chr(217)
    
    Print #filenum, String(33, "-")
    Print #filenum, Chr(27) & Chr(97) & Chr(48);
    Print #filenum, "PRESENTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "ACCEPTED BY:"
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, ""
    Print #filenum, "DATED:"
    Print #filenum, ""
    Print #filenum, Chr(27) & Chr(50);
    Print #filenum, String(33, "=")
    Print #filenum, Chr(27) & Chr(100) & Chr(7);
    Print #filenum, Chr(27) & Chr(64);
    Print #filenum, Chr(27) & Chr(69) & Chr(1);
    Close #1
    On Error GoTo 0
    frmTillReport.Tag = "   "
    Exit Sub
trap:
    If PrintErr = 0 Then
        PrintErr = 1
        For Each x In Printers
            If UCase(x.DeviceName) = UCase(Trim(Mid(Slip_Printer, (InStrRev(Slip_Printer, "\") + 1)))) Then
                Slip_Port = x.Port
                Exit For
            End If
        Next
        Resume Next
    End If
    Load frmError
    frmError.Caption = " Printer Error - " & Slip_Printer
    frmError.lblCap.Caption = "This Printer is currently Offline or not Installed. Please check your Printer Settings."
    frmError.lblError.Caption = err.Description
    DoEvents
    frmError.Show vbModal
    On Error GoTo 0
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 0
    If Val(frmCapture.Tag) = -10 Then
        lblHeading.Caption = Trim(UserRecord.User_Number) & " - " & Trim(UserRecord.Name) & " > Cashup No: " & frmCapture.lblHeading.Tag
        ActiveReadServer "Select * from Counters where Cashup_No = " & frmCapture.lblHeading.Tag
        lblRep.Caption = "-Float:"
        lblTot.Caption = "Total:"
        picFloat.Visible = True
    Else
        lblHeading.Caption = frmTillReport.cmdUser.Caption & " > Cashup No: " & frmTillReport.lblCashupInfo.Tag
        ActiveReadServer "Select * from Counters where Cashup_No = " & frmTillReport.lblCashupInfo.Tag
        lblRep.Caption = "Reported Total:"
        lblTot.Caption = "Variance:"
        picFloat.Visible = False
    End If
    If frmSplash.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        newBack.Visible = True
        For i = 0 To Me.Controls.Count - 1
            If Me.Controls(i).Name <> "newBack" Then
                Me.Controls(i).FontSize = Int(Me.Controls(i).FontSize * 0.65)
                Me.Controls(i).Width = Me.Controls(i).Width * 0.9
                Me.Controls(i).Left = Me.Controls(i).Left * 0.9
                Me.Controls(i).Height = Me.Controls(i).Height * 0.76
                Me.Controls(i).top = Me.Controls(i).top * 0.76
                Me.Controls(i).FontTextCaption.Size = Int(Me.Controls(i).FontTextCaption.Size * 0.9)
            End If
        Next i
        On Error GoTo 0
        newBack.Width = Me.Width
        newBack.Height = Me.Height
    End If
    If Val(frmCapture.Tag) <> -10 Then lblCount(1).Caption = Format(frmTillReport.lblCash(47).Caption, "0.00")
    frmCapture.Tag = ""
    If rs.Fields("Counted") <> 0 Then
        frmCapture.Tag = 0
        txtQty(0).Text = rs.Fields("200") & ""
        frmCapture.Tag = 1
        txtQty(1).Text = rs.Fields("100") & ""
        frmCapture.Tag = 2
        txtQty(2).Text = rs.Fields("50") & ""
        frmCapture.Tag = 3
        txtQty(3).Text = rs.Fields("20") & ""
        frmCapture.Tag = 4
        txtQty(4).Text = rs.Fields("10") & ""
        frmCapture.Tag = 5
        txtQty(5).Text = rs.Fields("5") & ""
        frmCapture.Tag = 6
        txtQty(6).Text = rs.Fields("2") & ""
        frmCapture.Tag = 7
        txtQty(7).Text = rs.Fields("1") & ""
        frmCapture.Tag = 8
        txtQty(8).Text = rs.Fields("50c") & ""
        frmCapture.Tag = 9
        txtQty(9).Text = rs.Fields("20c") & ""
        frmCapture.Tag = 10
        txtQty(10).Text = rs.Fields("10c") & ""
        frmCapture.Tag = 11
        txtQty(11).Text = rs.Fields("5c") & ""
        frmCapture.Tag = 12
        txtQty(12).Text = Format(rs.Fields("CardC") & "", "0.00")
        frmCapture.Tag = 13
        txtQty(13).Text = Format(rs.Fields("ChequeC") & "", "0.00")
        frmCapture.Tag = 14
        txtQty(14).Text = Format(rs.Fields("ChargeC") & "", "0.00")
        Calculate
    End If
   
    txtQty(0).SetFocus
    Calculate
    ' rs.Close
    If cmdKey(0).Enabled = False Then
        For i = 0 To 14
            txtQty(i).Locked = True
        Next i
    Else
        For i = 0 To 14
            txtQty(i).Locked = False
        Next i
    End If
    frmCapture.Tag = ""
End Sub
Private Sub Form_Load()
    If frmSplash.Height < 10000 And newBack.Visible = False Then
        On Error Resume Next
        DoEvents
        Me.Width = Me.Width * 0.9
        Me.Height = Me.Height * 0.9
        On Error GoTo 0
    End If
End Sub



Private Sub txtFloat_Change()
    Calculate
End Sub
Private Sub txtFloat_GotFocus()
    frmCapture.Tag = 16
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub
Private Sub txtFloat_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            txtFloat.Text = Format(txtFloat.Text, "0.00")
            txtQty(0).SetFocus
        Case 38
            KeyCode = 0
            txtFloat.Text = Format(txtFloat.Text, "0.00")
            txtQty(14).SetFocus
        Case 40
            KeyCode = 0
            txtFloat.Text = Format(txtFloat.Text, "0.00")
            txtQty(0).SetFocus
    End Select
End Sub
Private Sub txtFloat_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 46, 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
Private Sub txtQty_Change(Index As Integer)
    Calculate
End Sub
Private Sub txtQty_GotFocus(Index As Integer)
    On Error Resume Next
    If Val(frmCapture.Tag) = -1 Then frmCapture.Tag = Index
    Select Case Index
        Case 0 To 14
            If txtQty(Val(frmCapture.Tag)).Tag = "" Then
                ActiveControl.SelStart = 0
                ActiveControl.SelLength = Len(ActiveControl.Text)
            End If
            frmCapture.Tag = Index
    End Select
    On Error GoTo 0
End Sub
Private Sub txtQty_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If cmdKey(0).Enabled = False Then
        KeyCode = 0
        Exit Sub
    End If
    Select Case KeyCode
        Case 13
            Select Case Index
                Case 12: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
                Case 13: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
                Case 14: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
            End Select
            If Val(frmCapture.Tag) = 14 Then frmCapture.Tag = -1
            If Val(frmCapture.Tag) = -1 And lblTot.Caption = "Total:" Then
                txtFloat.SetFocus
                Exit Sub
            End If
            txtQty(Val(frmCapture.Tag) + 1).SetFocus
        Case 38
            Select Case Index
                Case 12: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
                Case 13: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
                Case 14: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
            End Select
            KeyCode = 0
            If Val(frmCapture.Tag) = 0 Then frmCapture.Tag = 15
            If Val(frmCapture.Tag) = 15 And lblTot.Caption = "Total:" Then
                txtFloat.SetFocus
                Exit Sub
            End If
            txtQty(Val(frmCapture.Tag) - 1).SetFocus
        Case 40
            Select Case Index
                Case 12: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
                Case 13: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
                Case 14: txtQty(Index).Text = Format(txtQty(Index).Text, "0.00")
            End Select
            KeyCode = 0
            If Val(frmCapture.Tag) = 14 Then frmCapture.Tag = -1
            If Val(frmCapture.Tag) = -1 And lblTot.Caption = "Total:" Then
                txtFloat.SetFocus
                Exit Sub
            End If
            txtQty(Val(frmCapture.Tag) + 1).SetFocus
    End Select
End Sub

Private Sub txtQty_KeyPress(Index As Integer, KeyAscii As Integer)
    If cmdKey(0).Enabled = False Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Index
        Case 0 To 11
            Select Case KeyAscii
                Case 8, 48 To 57
                Case Else: KeyAscii = 0
            End Select
        Case 12, 13, 14
            If InStr(ActiveControl.Text, ".") <> 0 And KeyAscii = 46 Then KeyAscii = 0
            Select Case KeyAscii
                Case 8, 46, 48 To 57
                Case Else: KeyAscii = 0
            End Select
    End Select
End Sub

