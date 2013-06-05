VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRes 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00CCFBFB&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   14235
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker mtView 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "ddd dd MMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   405
      Left            =   10230
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   180
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   16761024
      CustomFormat    =   "dddd dd MMM yyyy"
      Format          =   25755651
      CurrentDate     =   38895
   End
   Begin btButtonEx.ButtonEx cmdTop 
      Height          =   465
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Account"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin VSFlex8Ctl.VSFlexGrid grdRes 
      Height          =   8175
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   8325
      _cx             =   14684
      _cy             =   14420
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
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
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14417405
      BackColorAlternate=   16645618
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
      Rows            =   50
      Cols            =   34
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRes.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   540
         Top             =   630
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2850
         Top             =   1710
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   32
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":0277
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":0799
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":0CBB
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":11DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":16FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":1C21
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":2143
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":2665
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":2B87
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":30A9
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":35CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":3AED
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":400F
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":4531
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":4A53
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":4F75
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":5497
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":59B9
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":5EDB
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":63FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":691F
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":6E41
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":7363
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":7885
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":7DA7
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":82C9
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":87EB
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":8D0D
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":922F
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":9751
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":9BA3
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRes.frx":9FF5
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grdTask 
      Height          =   7890
      Left            =   8430
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   5235
      _cx             =   9234
      _cy             =   13917
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
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
      BackColorSel    =   16577005
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   600
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRes.frx":A447
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
   Begin btButtonEx.ButtonEx cmdTop 
      Height          =   465
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Groups"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdTop 
      Height          =   465
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Check In/Out"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdTop 
      Height          =   465
      Index           =   3
      Left            =   4440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Change"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdTop 
      Height          =   465
      Index           =   4
      Left            =   5880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Cancel"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdTop 
      Height          =   465
      Index           =   5
      Left            =   7320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Print"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdTask 
      Height          =   465
      Left            =   8760
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   150
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Show Tasks"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdLeft 
      Height          =   270
      Left            =   3450
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   " "
      Top             =   8970
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   476
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "3"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdRight 
      Height          =   270
      Left            =   7530
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8970
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   476
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "4"
      CaptionOffsetX  =   2
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdLeft1 
      Height          =   270
      Left            =   4050
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   " "
      Top             =   8970
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   476
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "7"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdRight1 
      Height          =   270
      Left            =   6930
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   " "
      Top             =   8970
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   476
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "8"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdCondition 
      Height          =   270
      Left            =   1680
      TabIndex        =   16
      Top             =   8970
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   476
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Conditions of Stay..."
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
   Begin btButtonEx.ButtonEx cmdRoomRates 
      Height          =   270
      Left            =   90
      TabIndex        =   17
      Top             =   8970
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   476
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Room Rates..."
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
   Begin btButtonEx.ButtonEx cmdAdd 
      Height          =   465
      Left            =   8430
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   820
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   12632256
      Caption         =   "New Task"
      Enabled         =   0   'False
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdDone 
      Height          =   465
      Left            =   10185
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   820
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   12632256
      Caption         =   "Mark as Completed"
      Enabled         =   0   'False
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdCanTask 
      Height          =   465
      Left            =   11940
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   820
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   12632256
      Caption         =   "Cancel Task"
      Enabled         =   0   'False
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0
   End
   Begin btButtonEx.ButtonEx cmdInfo 
      Height          =   270
      Left            =   4680
      TabIndex        =   22
      Top             =   8970
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   476
      Appearance      =   3
      BorderColor     =   12632256
      Caption         =   "Info..."
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
   Begin VB.Label lblType 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5910
      TabIndex        =   18
      Top             =   9000
      Width           =   2145
   End
   Begin VB.Label lblCol 
      Height          =   255
      Left            =   690
      TabIndex        =   15
      Top             =   9960
      Width           =   1215
   End
   Begin VB.Label lblRow 
      Height          =   315
      Left            =   690
      TabIndex        =   14
      Top             =   9600
      Width           =   1215
   End
   Begin MSForms.Image Image1 
      Height          =   345
      Left            =   60
      Top             =   8940
      Width           =   3330
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "5874;609"
   End
   Begin MSForms.Image Image3 
      Height          =   345
      Left            =   3420
      Top             =   8940
      Width           =   4980
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "8784;609"
   End
   Begin MSForms.Image Image2 
      Height          =   585
      Left            =   60
      Top             =   90
      Width           =   13605
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "23998;1032"
   End
   Begin MSForms.Image picImage 
      Height          =   9525
      Left            =   0
      Top             =   0
      Width           =   13755
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "24262;16801"
   End
End
Attribute VB_Name = "frmRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub FindMe(Res_No)
    rowStart = 0
    colStart = 0
    colStop = 0
    ActiveReadServer2 "Select * from Reservations where Res_No = " & Res_No
    If rs2.RecordCount > 0 Then
        mtView.Value = rs2.Fields("Arrive_Date")
        mtView_CloseUp
        For i = 3 To grdRes.Rows - 1
            If Val(rs2.Fields("Room_No") & "") = grdRes.ValueMatrix(i, 1) Then
                rowStart = i
                Exit For
            End If
        Next i
        For i = 2 To grdRes.Cols - 1
            If rs2.Fields("Arrive_Date") = grdRes.TextMatrix(2, i) Then
                colStart = i
            End If
            If rs2.Fields("Depart_Date") = grdRes.TextMatrix(2, i) Then
                colStop = i
            End If
        Next i
    End If
    If colStop = 0 Then colStop = 33
    rs2.Close
    grdRes.FocusRect = flexFocusNone
    Timer1.Enabled = True
    grdRes.ToolTipText = ""
End Sub
Private Sub cmdadd_Click()
    frmTask.Show vbModal
    LoadTasks
End Sub
Private Sub cmdCanTask_Click()
    Response = MsgBox("Are you sure you want to cancel this task?", vbYesNo, "Room Tasks")
    If Response = vbYes Then
        ActiveUpdateServer " Delete from Room_Tasks where Res_No = " & TillData.Res_No & " and Task_No = " & grdTask.TextMatrix(grdTask.Row, 2)
        LoadTasks
    End If
End Sub

Private Sub cmdCondition_Click()
    frmConditions.Show vbModal
End Sub
Private Sub cmdDone_Click()
    Response = MsgBox("Are you sure this task is completed?", vbYesNo, "Room Tasks")
    If Response = vbYes Then
        ActiveUpdateServer " Update Room_Tasks set Completed = 1 where Res_No = " & TillData.Res_No & " and Task_No = " & grdTask.TextMatrix(grdTask.Row, 2)
        DoEvents
        Set grdTask.CellPicture = ImageList1.ListImages(31).Picture
        grdTask.CellPictureAlignment = flexPicAlignCenterCenter
        cmdDone.Enabled = False
        Cmdadd.Enabled = True
        LoadTasks
    End If
End Sub
Private Sub cmdInfo_Click()
    frmInfo.Show vbModal
End Sub
Private Sub cmdLeft_Click()
    For b = 3 To grdRes.Rows - 1
        grdRes.Cell(flexcpBackColor, b, 2, b, 2) = grdRes.Cell(flexcpBackColor, b, 33, b, 33)
    Next b
    For i = 32 To 3 Step -1
        For b = 3 To grdRes.Rows - 1
            grdRes.Cell(flexcpBackColor, b, i, b, i) = grdRes.Cell(flexcpBackColor, b, 33, b, 33)
        Next b
        grdRes.TextMatrix(2, i) = DateAdd("D", -1, grdRes.TextMatrix(2, i))
        grdRes.TextMatrix(0, i) = Format(grdRes.TextMatrix(2, i), "D")
        Select Case Format(grdRes.TextMatrix(2, i), "w")
            Case 1
                grdRes.TextMatrix(1, i) = "Su"
                grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
            Case 2: grdRes.TextMatrix(1, i) = "M"
            Case 3: grdRes.TextMatrix(1, i) = "T"
            Case 4: grdRes.TextMatrix(1, i) = "W"
            Case 5: grdRes.TextMatrix(1, i) = "Th"
            Case 6: grdRes.TextMatrix(1, i) = "F"
            Case 7
                grdRes.TextMatrix(1, i) = "Sa"
                grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
        End Select
    Next i
    grdRes.TextMatrix(2, i) = DateAdd("D", -1, grdRes.TextMatrix(2, i))
    grdRes.TextMatrix(0, i) = Format(grdRes.TextMatrix(2, i), "D")
    Select Case Format(grdRes.TextMatrix(2, i), "w")
        Case 1
            grdRes.TextMatrix(1, i) = "Su"
            grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
        Case 2: grdRes.TextMatrix(1, i) = "M"
        Case 3: grdRes.TextMatrix(1, i) = "T"
        Case 4: grdRes.TextMatrix(1, i) = "W"
        Case 5: grdRes.TextMatrix(1, i) = "Th"
        Case 6: grdRes.TextMatrix(1, i) = "F"
        Case 7
            grdRes.TextMatrix(1, i) = "Sa"
            grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
    End Select
    If Format(mtView.Value, "YYYY-MM-DD") > Format(grdRes.TextMatrix(2, 32), "YYYY-MM-DD") Then
        mtView.Value = Format(grdRes.TextMatrix(2, 32), "M/DD/YYYY")
    End If
    For i = 2 To 33
        If Format(mtView.Value, "YYYY-MM-DD") = Format(grdRes.TextMatrix(2, i), "YYYY-MM-DD") Then
            grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HFDE1D7
        End If
    Next i
    
End Sub

Private Sub cmdLeft_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    LoadRes
End Sub

Private Sub cmdLeft1_Click()
    For i = 1 To 7
         cmdLeft_Click
    Next i
    LoadRes
End Sub

Private Sub cmdRight_Click()
    For b = 3 To grdRes.Rows - 1
        grdRes.Cell(flexcpBackColor, b, 32, b, 32) = grdRes.Cell(flexcpBackColor, b, 33, b, 33)
    Next b
    For i = 2 To 31
        grdRes.TextMatrix(2, i) = DateAdd("D", 1, grdRes.TextMatrix(2, i))
        grdRes.TextMatrix(0, i) = Format(grdRes.TextMatrix(2, i), "D")
        For b = 1 To grdRes.Rows - 1
            grdRes.Cell(flexcpBackColor, b, i, b, i) = grdRes.Cell(flexcpBackColor, b, 33, b, 33)
        Next b
        Select Case Format(grdRes.TextMatrix(2, i), "w")
            Case 1
                grdRes.TextMatrix(1, i) = "Su"
                grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
            Case 2: grdRes.TextMatrix(1, i) = "M"
            Case 3: grdRes.TextMatrix(1, i) = "T"
            Case 4: grdRes.TextMatrix(1, i) = "W"
            Case 5: grdRes.TextMatrix(1, i) = "Th"
            Case 6: grdRes.TextMatrix(1, i) = "F"
            Case 7
                grdRes.TextMatrix(1, i) = "Sa"
                grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
        End Select
    Next i
    grdRes.TextMatrix(2, i) = DateAdd("D", 1, grdRes.TextMatrix(2, i))
    grdRes.TextMatrix(0, i) = Format(grdRes.TextMatrix(2, i), "D")
    Select Case Format(grdRes.TextMatrix(2, i), "w")
        Case 1
            grdRes.TextMatrix(1, i) = "Su"
            grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
        Case 2: grdRes.TextMatrix(1, i) = "M"
        Case 3: grdRes.TextMatrix(1, i) = "T"
        Case 4: grdRes.TextMatrix(1, i) = "W"
        Case 5: grdRes.TextMatrix(1, i) = "Th"
        Case 6: grdRes.TextMatrix(1, i) = "F"
        Case 7
            grdRes.TextMatrix(1, i) = "Sa"
            grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
    End Select
    
    If Format(mtView.Value, "YYYY-MM-DD") < Format(grdRes.TextMatrix(2, 2), "YYYY-MM-DD") Then
        mtView.Value = Format(grdRes.TextMatrix(2, 2), "M/DD/YYYY")
    End If
    For i = 2 To 33
        If Format(mtView.Value, "YYYY-MM-DD") = Format(grdRes.TextMatrix(2, i), "YYYY-MM-DD") Then
            grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HFDE1D7
        End If
    Next i
    
End Sub
Private Sub cmdRight_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    LoadRes
End Sub

Private Sub cmdRight1_Click()
    For i = 1 To 7
         cmdRight_Click
    Next i
    LoadRes
End Sub
Private Sub cmdTask_Click()
    Select Case cmdTask.Caption
        Case "Show Tasks"
            grdRes.Width = 8325
            cmdTask.Caption = "Hide Tasks"
            grdTask.Visible = True
            Cmdadd.Visible = True
            cmdDone.Visible = True
            cmdCanTask.Visible = True
            cmdTask.BackColor = &HFFC0C0
            cmdTask.FontBold = True
        Case "Hide Tasks"
            grdRes.Width = 13605
            grdTask.Visible = False
            cmdTask.Caption = "Show Tasks"
            Cmdadd.Visible = False
            cmdDone.Visible = False
            cmdCanTask.Visible = False
            cmdTask.BackColor = &H8000000F
            cmdTask.FontBold = False
    End Select
    Image1.Width = grdRes.ColWidth(0) + grdRes.ColWidth(1)
    Image3.Left = grdRes.ColWidth(0) + grdRes.ColWidth(1) + 90
    Image3.Width = grdRes.Width - (grdRes.ColWidth(0) + grdRes.ColWidth(1) + 30)
    cmdLeft.Left = Image3.Left + 50
    cmdRight.Left = Image3.Left + Image3.Width - 640
    cmdLeft1.Left = cmdLeft.Left + 600
    cmdRight1.Left = cmdRight.Left - 600
End Sub

Private Sub cmdTop_Click(Index As Integer)
    If Index = 1 Then
        
        
        Exit Sub
    End If
    For i = 0 To cmdTop.Count - 1
        If i = Index Then
            cmdTop(i).BackColor = &HFFC0C0
            cmdTop(i).FontBold = True
            If cmdTop(i).Value = Up Then
                cmdTop(i).BackColor = &H8000000F
                cmdTop(i).FontBold = False
            End If
        Else
            cmdTop(i).Value = Up
            cmdTop(i).BackColor = &H8000000F
            cmdTop(i).FontBold = False
        End If
    Next i
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            KeyCode = 0
            ActiveUpdateServer "Insert into User_Journal (User_No,Date_Time,Function_Key,Workstation_No) values (" & UserRecord.User_Number & ",Getdate(),2," & Workstation_No & ")"
            frmSplash.Show
            frmMain.Hide
    End Select
End Sub
Private Sub Form_Load()
    grdTask.Cols = 4
    grdTask.ColHidden(2) = True
    grdTask.ColHidden(3) = True
    Cmdadd.Visible = False
    cmdDone.Visible = False
    cmdCanTask.Visible = False
    grdRes.Rows = 3
    grdRes.RowHidden(2) = True
    grdRes.ColHidden(33) = True
    grdRes.Width = 13605
    grdTask.Visible = False
    cmdTask.Caption = "Show Tasks"
    grdRes.ColWidth(0) = 2750
    grdRes.ColWidth(1) = 600
    grdRes.TextMatrix(0, 0) = "Room Description"
    grdRes.TextMatrix(0, 1) = "No."
    grdRes.TextMatrix(1, 0) = "Room Description"
    grdRes.TextMatrix(1, 1) = "No."
    grdRes.MergeCells = flexMergeFixedOnly
    grdRes.MergeCol(0) = True
    grdRes.MergeCol(1) = True
    grdRes.ColAlignment(0) = flexAlignLeftCenter
    grdRes.ColAlignment(1) = flexAlignLeftCenter
    ActiveReadServer "select * from Rooms order by convert(int,replace(replace(replace(room_No,'a',''),'b',''),'c',''))"
    While Not rs.EOF
        grdRes.Rows = grdRes.Rows + 1
        grdRes.TextMatrix(grdRes.Rows - 1, 0) = rs.Fields("Description")
        grdRes.TextMatrix(grdRes.Rows - 1, 1) = rs.Fields("Room_No")
        rs.MoveNext
    Wend
    rs.Close
    grdRes.Cell(flexcpBackColor, 2, 0, grdRes.Rows - 1, 0) = &HFEF0EB
    grdRes.Cell(flexcpBackColor, 2, 1, grdRes.Rows - 1, 1) = &HDFFCFD
    mtView.Value = Date
    Select Case mtView.Month
        Case 1: Lastday = 31
        Case 2
            If Int(mtView.Year / 4) = (mtView.Year / 4) Then
                Lastday = 29
            Else
                Lastday = 28
            End If
        Case 3: Lastday = 31
        Case 4: Lastday = 30
        Case 5: Lastday = 31
        Case 6: Lastday = 30
        Case 7: Lastday = 31
        Case 8: Lastday = 31
        Case 9: Lastday = 30
        Case 10: Lastday = 31
        Case 11: Lastday = 30
        Case 12: Lastday = 31
    End Select
    CurrentDate = mtView.Value
    For i = 2 To 32
        grdRes.ColWidth(i) = 330
        If i - 1 = mtView.Day Then
            Marker = i
        End If
        If i - 1 > mtView.Day Then
            grdRes.TextMatrix(2, i) = DateAdd("d", 1, CurrentDate)
            CurrentDate = DateAdd("d", 1, CurrentDate)
            grdRes.TextMatrix(0, i) = Format(CurrentDate, "D")
            Select Case Format(CurrentDate, "w")
                Case 1
                    grdRes.TextMatrix(1, i) = "Su"
                    grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
                Case 2: grdRes.TextMatrix(1, i) = "M"
                Case 3: grdRes.TextMatrix(1, i) = "T"
                Case 4: grdRes.TextMatrix(1, i) = "W"
                Case 5: grdRes.TextMatrix(1, i) = "Th"
                Case 6: grdRes.TextMatrix(1, i) = "F"
                Case 7
                    grdRes.TextMatrix(1, i) = "Sa"
                    grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
            End Select
        End If
        grdRes.ColAlignment(i) = flexAlignCenterCenter
    Next i
    If Marker = 32 Then
        grdRes.TextMatrix(0, 32) = 31
        grdRes.TextMatrix(2, 32) = mtView.Value
    End If
    For i = Marker To 2 Step -1
        If i <> 32 Then
            grdRes.TextMatrix(0, i) = Format(DateAdd("d", -1, grdRes.TextMatrix(2, i + 1)), "D")
            grdRes.TextMatrix(2, i) = DateAdd("d", -1, grdRes.TextMatrix(2, i + 1))
        End If
        Select Case Format(grdRes.TextMatrix(2, i), "w")
            Case 1
                grdRes.TextMatrix(1, i) = "Su"
                grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
            Case 2: grdRes.TextMatrix(1, i) = "M"
            Case 3: grdRes.TextMatrix(1, i) = "T"
            Case 4: grdRes.TextMatrix(1, i) = "W"
            Case 5: grdRes.TextMatrix(1, i) = "Th"
            Case 6: grdRes.TextMatrix(1, i) = "F"
            Case 7
                grdRes.TextMatrix(1, i) = "Sa"
                grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
        End Select
    Next i
    LoadRes
    grdRes.Cell(flexcpBackColor, 2, Marker, grdRes.Rows - 1, Marker) = &HFDE1D7
    If grdRes.Rows > 28 Then
        grdRes.ColWidth(0) = 2450
    End If
    Image1.Width = grdRes.ColWidth(0) + grdRes.ColWidth(1)
    Image3.Left = grdRes.ColWidth(0) + grdRes.ColWidth(1) + 90
    Image3.Width = grdRes.Width - (grdRes.ColWidth(0) + grdRes.ColWidth(1) + 30)
    cmdLeft.Left = Image3.Left + 50
    cmdRight.Left = Image3.Left + Image3.Width - 640
    cmdLeft1.Left = cmdLeft.Left + 600
    cmdRight1.Left = cmdRight.Left - 600
End Sub
Public Sub LoadRes()
    DoEvents
    ActiveReadServer "Select Arrive_Date,Depart_Date,Res_No,Room_No,Res_Type,Title + ' ' + First_Name + ' ' + Last_name as Guest_Name,Res_Type from Reservations " & _
    " where (Arrive_Date > '" & DateAdd("D", -1, grdRes.TextMatrix(2, 2)) & "'and Arrive_Date < '" & DateAdd("D", 1, grdRes.TextMatrix(2, 32)) & "')" & _
    " or (Depart_Date > '" & DateAdd("D", -1, grdRes.TextMatrix(2, 2)) & "'and Depart_Date < '" & DateAdd("D", 1, grdRes.TextMatrix(2, 32)) & "')" & _
    " order by Room_No, Line_No"
    For i = 3 To grdRes.Rows - 1
        For ib = 2 To 32
            grdRes.Select i, ib, i, ib
            Set grdRes.CellPicture = LoadPicture("")
            grdRes.TextMatrix(i, ib) = ""
        Next ib
    Next i
    While Not rs.EOF
        For i = 3 To grdRes.Rows - 1
            If Val(grdRes.TextMatrix(i, 1)) = Val(rs.Fields("Room_No")) Then
                grdRes.Row = i
                Exit For
            End If
        Next i
        If Format(rs.Fields("Arrive_Date"), vbGeneralDate) < Format(grdRes.TextMatrix(2, 2), vbGeneralDate) Then
            For b = 2 To 32
                If grdRes.TextMatrix(2, b) <> rs.Fields("Depart_Date") Then
                    grdRes.TextMatrix(grdRes.Row, b) = rs.Fields("Res_Type") & "< M: " & rs.Fields("Res_No") & " - " & rs.Fields("Guest_Name")
                    Load_Icon grdRes.TextMatrix(grdRes.Row, b), grdRes.Row, b
                Else
                    grdRes.TextMatrix(grdRes.Row, b) = grdRes.TextMatrix(grdRes.Row, b) & "|" & rs.Fields("Res_Type") & "< D: " & rs.Fields("Res_No") & " - " & rs.Fields("Guest_Name")
                    Load_Icon grdRes.TextMatrix(grdRes.Row, b), grdRes.Row, b
                    Exit For
                End If
            Next b
        Else
            For i = 2 To 32
                If rs.Fields("Arrive_Date") = grdRes.TextMatrix(2, i) Then
                    Exit For
                End If
            Next i
            If InStr(grdRes.TextMatrix(grdRes.Row, i), "D") <> 0 Then
                grdRes.TextMatrix(grdRes.Row, i) = rs.Fields("Res_Type") & "< A: " & rs.Fields("Res_No") & " - " & rs.Fields("Guest_Name") & grdRes.TextMatrix(grdRes.Row, i)
            Else
                grdRes.TextMatrix(grdRes.Row, i) = grdRes.TextMatrix(grdRes.Row, i) & "|" & rs.Fields("Res_Type") & "< A: " & rs.Fields("Res_No") & " - " & rs.Fields("Guest_Name")
            End If
            Load_Icon grdRes.TextMatrix(grdRes.Row, i), grdRes.Row, i
            For b = i + 1 To 32
                If grdRes.TextMatrix(2, b) <> rs.Fields("Depart_Date") Then
                    grdRes.TextMatrix(grdRes.Row, b) = rs.Fields("Res_Type") & "< M: " & rs.Fields("Res_No") & " - " & rs.Fields("Guest_Name")
                    Load_Icon grdRes.TextMatrix(grdRes.Row, b), grdRes.Row, b
                Else
                    grdRes.TextMatrix(grdRes.Row, b) = grdRes.TextMatrix(grdRes.Row, b) & "|" & rs.Fields("Res_Type") & "< D: " & rs.Fields("Res_No") & " - " & rs.Fields("Guest_Name")
                    Load_Icon grdRes.TextMatrix(grdRes.Row, b), grdRes.Row, b
                    Exit For
                End If
            Next b
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub Load_Icon(IconString, Row, Col)
    grdRes.Select Row, Col, Row, Col
    If InStr(IconString, "M:") <> 0 Then
        Res_Type = Val(Replace(Mid(IconString, 1, InStr(IconString, "<")), "|", ""))
        Select Case Res_Type
            Case 0
                Set grdRes.CellPicture = ImageList1.ListImages(8).Picture
            Case 1
                Set grdRes.CellPicture = ImageList1.ListImages(6).Picture
            Case 2
                Set grdRes.CellPicture = ImageList1.ListImages(5).Picture
            Case 3
                Set grdRes.CellPicture = ImageList1.ListImages(7).Picture
        End Select
    Else
        If InStr(IconString, "A:") <> 0 And InStr(IconString, "D:") <> 0 Then
            StringA = Mid(IconString, InStrRev(IconString, "|"))
            stringD = Mid(IconString, 1, InStrRev(IconString, "|") - 1)
            If InStr(StringA, "D") = 0 Then
                SaveString = StringA
                StringA = stringD
                stringD = SaveString
            End If
            If InStr(stringD, "A") = 0 And stringD <> "" Then
                SaveString = StringA
                StringA = stringD
                stringD = SaveString
            End If
            Res_TypeA = Val(Replace(Mid(StringA, 1, InStr(StringA, "<")), "|", ""))
            Res_TypeD = Val(Replace(Mid(stringD, 1, InStr(stringD, "<")), "|", ""))
            Select Case Trim(Str(Res_TypeA)) & "-" & Trim(Str(Res_TypeD))
                Case "0-0": Set grdRes.CellPicture = ImageList1.ListImages(26).Picture
                Case "0-1": Set grdRes.CellPicture = ImageList1.ListImages(28).Picture
                Case "0-2": Set grdRes.CellPicture = ImageList1.ListImages(27).Picture
                Case "0-3": Set grdRes.CellPicture = ImageList1.ListImages(29).Picture
                Case "1-0": Set grdRes.CellPicture = ImageList1.ListImages(21).Picture
                Case "1-1": Set grdRes.CellPicture = ImageList1.ListImages(18).Picture
                Case "1-2": Set grdRes.CellPicture = ImageList1.ListImages(19).Picture
                Case "1-3": Set grdRes.CellPicture = ImageList1.ListImages(20).Picture
                Case "2-0": Set grdRes.CellPicture = ImageList1.ListImages(17).Picture
                Case "2-1": Set grdRes.CellPicture = ImageList1.ListImages(15).Picture
                Case "2-2": Set grdRes.CellPicture = ImageList1.ListImages(14).Picture
                Case "2-3": Set grdRes.CellPicture = ImageList1.ListImages(16).Picture
                Case "3-0": Set grdRes.CellPicture = ImageList1.ListImages(25).Picture
                Case "3-1": Set grdRes.CellPicture = ImageList1.ListImages(24).Picture
                Case "3-2": Set grdRes.CellPicture = ImageList1.ListImages(23).Picture
                Case "3-3": Set grdRes.CellPicture = ImageList1.ListImages(22).Picture
            End Select
        Else
            If InStr(IconString, "A:") <> 0 Then
                Res_Type = Val(Replace(Mid(IconString, 1, InStr(IconString, "<")), "|", ""))
                Select Case Res_Type
                    Case 0
                        Set grdRes.CellPicture = ImageList1.ListImages(4).Picture
                    Case 1
                        Set grdRes.CellPicture = ImageList1.ListImages(2).Picture
                    Case 2
                        Set grdRes.CellPicture = ImageList1.ListImages(1).Picture
                    Case 3
                        Set grdRes.CellPicture = ImageList1.ListImages(3).Picture
                End Select
            End If
            If InStr(IconString, "D:") <> 0 Then
                Res_Type = Val(Replace(Mid(IconString, 1, InStr(IconString, "<")), "|", ""))
                Select Case Res_Type
                    Case 0
                        Set grdRes.CellPicture = ImageList1.ListImages(12).Picture
                    Case 1
                        Set grdRes.CellPicture = ImageList1.ListImages(10).Picture
                    Case 2
                        Set grdRes.CellPicture = ImageList1.ListImages(9).Picture
                    Case 3
                        Set grdRes.CellPicture = ImageList1.ListImages(11).Picture
                End Select
            End If
        End If
    End If
End Sub
Private Sub LoadTasks()
    On Error Resume Next
    cmdDone.Enabled = False
    If grdRes.TextMatrix(grdRes.Row, grdRes.Col) <> "" Then
        grdTask.TextMatrix(0, 1) = ""
        grdTask.Rows = 1
        TillData.Res_No = Val(Mid(grdRes.TextMatrix(grdRes.Row, grdRes.Col), InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1, (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "-") - 1) - (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1)))
        ActiveReadServer "Select * from Reservations where Res_no = " & Val(TillData.Res_No)
        If rs.RecordCount > 0 Then
            grdTask.TextMatrix(0, 1) = "All Task for Room: " & rs.Fields("Room_No") & " - " & rs.Fields("Title") & " " & rs.Fields("First_Name") & " " & rs.Fields("Last_Name")
            If Trim(rs.Fields("Remarks") & "") <> "" Then
                grdTask.Rows = grdTask.Rows + 1
                grdTask.TextMatrix(grdTask.Rows - 1, 1) = rs.Fields("Remarks")
                Rows = Round(Len(rs.Fields("Remarks")) / 35, 0)
                If Rows = 1 Or Rows = 0 Then Rows = 2
                grdTask.RowHeight(grdTask.Rows - 1) = grdTask.RowHeight(0) * Rows
                grdTask.Select grdTask.Rows - 1, 0, grdTask.Row, 0
                Set grdTask.CellPicture = ImageList1.ListImages(32).Picture
                grdTask.CellPictureAlignment = flexPicAlignCenterCenter
                grdTask.TextMatrix(1, 3) = "2"
                cmdCanTask.Enabled = False
                cmdDone.Enabled = False
            End If
        End If
        rs.Close
        ActiveReadServer "Select * from Room_Tasks where Res_No=" & Val(TillData.Res_No)
        While Not rs.EOF
            If Trim(rs.Fields("Remarks") & "") <> Trim(grdTask.TextMatrix(1, 1)) Then
                grdTask.Rows = grdTask.Rows + 1
            End If
            grdTask.TextMatrix(grdTask.Rows - 1, 0) = ""
            grdTask.TextMatrix(grdTask.Rows - 1, 1) = rs.Fields("Description") & " - " & rs.Fields("Remarks")
            Rows = Round(Len(rs.Fields("Description") & " - " & rs.Fields("Remarks")) / 35, 0)
            If Rows = 1 Or Rows = 0 Then Rows = 2
            grdTask.RowHeight(grdTask.Rows - 1) = grdTask.RowHeight(0) * Rows
            Select Case Val(rs.Fields("Completed") & "")
                Case 0
                    grdTask.Select grdTask.Rows - 1, 0, grdTask.Row, 0
                    Set grdTask.CellPicture = ImageList1.ListImages(30).Picture
                    grdTask.CellPictureAlignment = flexPicAlignCenterCenter
                    cmdDone.Enabled = True
                Case 1
                    grdTask.Select grdTask.Rows - 1, 0, grdTask.Row, 0
                    Set grdTask.CellPicture = ImageList1.ListImages(31).Picture
                    grdTask.CellPictureAlignment = flexPicAlignCenterCenter
                    cmdDone.Enabled = False
            End Select
            grdTask.TextMatrix(grdTask.Rows - 1, 2) = rs.Fields("Task_no")
            grdTask.TextMatrix(grdTask.Rows - 1, 3) = Val(rs.Fields("Completed") & "")
            rs.MoveNext
        Wend
        rs.Close
        Cmdadd.Enabled = True
    Else
        Cmdadd.Enabled = False
        
        cmdCanTask.Enabled = False
    End If
    grdTask.AutoSizeMode = flexAutoSizeRowHeight
    grdTask.WordWrap = True
    If grdTask.Rows > 1 Then grdTask.Row = 1
    On Error GoTo 0
End Sub
Private Sub grdRes_Click()
    grdRes.FocusRect = flexFocusSolid
    Timer1.Enabled = False
    grdTask.Rows = 1
    grdTask.TextMatrix(0, 1) = ""
    Oneclick = False
    For i = 0 To cmdTop.Count - 1
        If cmdTop(i).Value = Down Then
            Oneclick = True
        End If
    Next i
    LoadTasks
    If grdRes.TextMatrix(grdRes.Row, grdRes.Col) = "" Then Exit Sub
    If Oneclick = False Then Exit Sub
    If InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "A:") = 0 And InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "D:") <> 0 Then
        TillData.Res_No = Val(Mid(grdRes.TextMatrix(grdRes.Row, grdRes.Col), InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1, (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "-") - 1) - (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1)))
    End If
    If InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "D:") <> 0 And InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "A:") <> 0 Then
        Load frmChooseRes
        frmChooseRes.Caption = "Select a Reservation"
        StringA = Mid(grdRes.TextMatrix(grdRes.Row, grdRes.Col), InStrRev(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "|"))
        stringD = Mid(grdRes.TextMatrix(grdRes.Row, grdRes.Col), 1, InStrRev(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "|") - 1)
        If InStr(StringA, "D") = 0 Then
            SaveString = StringA
            StringA = stringD
            stringD = SaveString
        End If
        Res_TypeA = Val(Replace(Mid(StringA, 1, InStr(StringA, "<")), "|", ""))
        Res_No = Val(Mid(StringA, InStr(StringA, ":") + 1, (InStr(StringA, "-") - 1) - (InStr(StringA, ":") + 1)))
        frmChooseRes.cmdRes(0).Caption = Trim(Mid(StringA, InStrRev(StringA, "-") + 1))
        frmChooseRes.cmdRes(0).Tag = Res_No
        Select Case Res_TypeA
            Case 0: frmChooseRes.cmdRes(0).BackColor = &HC0FFFF
            Case 1: frmChooseRes.cmdRes(0).BackColor = &HC0FFC0
            Case 2: frmChooseRes.cmdRes(0).BackColor = &HFFC0C0
            Case 3: frmChooseRes.cmdRes(0).BackColor = &HC0C0FF
        End Select
        Res_TypeD = Val(Replace(Mid(stringD, 1, InStr(stringD, "<")), "|", ""))
        Res_No = Val(Mid(stringD, InStr(stringD, ":") + 1, (InStr(stringD, "-") - 1) - (InStr(stringD, ":") + 1)))
        frmChooseRes.cmdRes(1).Caption = Trim(Mid(stringD, InStrRev(stringD, "-") + 1))
        frmChooseRes.cmdRes(1).Tag = Res_No
        Select Case Res_TypeD
            Case 0: frmChooseRes.cmdRes(1).BackColor = &HC0FFFF
            Case 1: frmChooseRes.cmdRes(1).BackColor = &HC0FFC0
            Case 2: frmChooseRes.cmdRes(1).BackColor = &HFFC0C0
            Case 3: frmChooseRes.cmdRes(1).BackColor = &HC0C0FF
        End Select
        frmChooseRes.Show vbModal
        If Val(frmChooseRes.Tag) = 0 Then
            Unload frmChooseRes
            Exit Sub
        End If
        TillData.Res_No = frmChooseRes.Tag
        Unload frmChooseRes
    Else
        If InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "A:") <> 0 Then
            TillData.Res_No = Val(Mid(grdRes.TextMatrix(grdRes.Row, grdRes.Col), InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1, (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "-") - 1) - (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1)))
        End If
        If InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "D:") <> 0 Then
            TillData.Res_No = Val(Mid(grdRes.TextMatrix(grdRes.Row, grdRes.Col), InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1, (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "-") - 1) - (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1)))
        End If
        If InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "D:") = 0 And InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "A:") = 0 Then
            TillData.Res_No = Val(Mid(grdRes.TextMatrix(grdRes.Row, grdRes.Col), InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1, (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "-") - 1) - (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1)))
        End If
    End If
    For i = 0 To cmdTop.Count - 1
        If cmdTop(i).Value = Down Then
            Select Case cmdTop(i).Caption
                Case "Groups"
                    
                Case "Account"
                    ActiveReadServer "Select Res_Type from Res_View where Res_No = " & TillData.Res_No & " group by Res_No, Res_Type"
                    If rs.RecordCount > 0 Then
                        Select Case rs.Fields("Res_Type")
                            Case 0
                                MsgBox "There is no Room Account open for this Reservation."
                            Case 1, 2, 3
                                rs.Close
                                Load frmAccount
                                frmAccount.Tag = "Room"
                                frmRes.Tag = ""
                                frmAccount.Show vbModal
                                Exit Sub
                        End Select
                    End If
                    rs.Close
                Case "Check In/Out"
                    ActiveReadServer "Select Res_Type from Res_View where Res_No = " & TillData.Res_No & " group by Res_No, Res_Type"
                    If rs.RecordCount > 0 Then
                        Select Case rs.Fields("Res_Type")
                            Case 0, 1, 2
                                frmRes.Tag = ""
                                rs.Close
                                frmCheckin.Show vbModal
                                If frmRes.Tag = "2" Then
                                    LoadRes
                                    ActiveReadServer "Select max(Invoice_No) as Invoice_No from Sales_Journal where Res_no = " & TillData.Res_No
                                    If rs.RecordCount > 0 Then
                                        TillData.Res_No = rs.Fields("Invoice_No")
                                    End If
                                    rs.Close
                                    rptInvoice.Show
                                End If
                                If frmRes.Tag = "3" Then
                                    LoadRes
                                    rptRoomAcc.Show
                                End If
                                Exit Sub
                            Case 3
                                Exit Sub
                        End Select
                    End If
                    rs.Close
                Case "Change"
                    ActiveReadServer "Select Res_Type from Res_View where Res_No = " & TillData.Res_No & " group by Res_No,Res_Type"
                    If rs.RecordCount > 0 Then
                        Select Case rs.Fields("Res_Type")
                            Case 0, 1, 2
                                rs.Close
                                frmRes.Tag = ""
                                frmCheckin.Show vbModal
                                If frmRes.Tag = "4" Then
                                    LoadRes
                                    ActiveReadServer "Select max(Invoice_No) as Invoice_No from Sales_Journal where Res_no = " & TillData.Res_No
                                    If rs.RecordCount > 0 Then
                                        TillData.Res_No = rs.Fields("Invoice_No")
                                    End If
                                    rs.Close
                                    rptInvoice.Show
                                    Exit Sub
                                End If
                                LoadRes
                                Exit Sub
                            Case 3
                                MsgBox "You cannot Change a Checked Out Reservation.", vbInformation, "HeroPOS"
                        End Select
                    End If
                    rs.Close
                Case "Print"
                    ActiveReadServer "Select Res_Type from Res_View where Res_No = " & TillData.Res_No & " group by Res_No, Res_Type"
                    If rs.RecordCount > 0 Then
                        Select Case rs.Fields("Res_Type")
                            Case 0
                                rs.Close
                                rptCheckin.Show
                                Exit Sub
                            Case 1
                                rs.Close
                                rptCheckin.Show
                                Exit Sub
                            Case 2, 3
                                rs.Close
                                rptRoomAcc.Show
                                Exit Sub
                        End Select
                    End If
                    rs.Close
                Case "Cancel"
                    ActiveReadServer "Select Res_Type,Balance from Res_View where Res_No = " & TillData.Res_No & " group by Res_No, Res_Type,Balance"
                    If rs.RecordCount > 0 Then
                        If Val(rs.Fields("Balance") & "") <> 0 Then
                            Select Case rs.Fields("Res_Type")
                                Case 0
                                    Response = MsgBox("Are you sure you want to Cancel this Reservation?", vbYesNo, "HeroPOS")
                                    If Response = vbYes Then
                                        ActiveUpdateServer "Delete from Reservations where Res_No=" & TillData.Res_No
                                        rs.Close
                                        LoadRes
                                        Exit Sub
                                    End If
                                Case 1
                                    MsgBox "You cannot Cancel this Reservation." & Chr(13) & "You can only cancel a Reservation if the Account Balance is Zero.", vbInformation, "HeroPOS"
                                Case 2
                                    MsgBox "You cannot Cancel this Reservation." & Chr(13) & "You have to Check the Guest Out.", vbInformation, "HeroPOS"
                                Case 3
                                    MsgBox "You cannot Cancel this Reservation.", vbInformation, "HeroPOS"
                            End Select
                        Else
                            Select Case rs.Fields("Res_Type")
                                Case 0
                                    Response = MsgBox("Are you sure you want to Cancel this Reservation?", vbYesNo, "HeroPOS")
                                    If Response = vbYes Then
                                        ActiveUpdateServer "Delete from Reservations where Res_No=" & TillData.Res_No
                                        rs.Close
                                        LoadRes
                                        Exit Sub
                                    End If
                                Case 1:
                                    Response = MsgBox("Are you sure you want to Cancel this Reservation?", vbYesNo, "HeroPOS")
                                    If Response = vbYes Then
                                        ActiveUpdateServer "Delete from Reservations where Res_No=" & TillData.Res_No
                                        rs.Close
                                        LoadRes
                                        Exit Sub
                                    End If
                                Case 2
                                    MsgBox "You cannot Cancel this Reservation." & Chr(13) & "You have to Check the Guest Out.", vbInformation, "HeroPOS"
                                Case 3:
                                    MsgBox "You cannot Cancel this Reservation.", vbInformation, "HeroPOS"
                            End Select
                        End If
                    End If
                    rs.Close
            End Select
        End If
    Next i
End Sub

Private Sub grdRes_DblClick()
    On Error GoTo trap
    Oneclick = False
    For i = 0 To cmdTop.Count - 1
        If cmdTop(i).Value = Down Then
            Oneclick = True
        End If
    Next i
    If Oneclick = True Then Exit Sub
    frmRes.Tag = ""
    SaveRow = grdRes.Row
    savecol = grdRes.Col
    frmCheckin.Show vbModal
    On Error GoTo 0
    If frmRes.Tag = "1" Then
        frmRes.Tag = ""
        LoadRes
        Response = MsgBox("Do you want to Print the Check In Form now?", vbYesNo, "HeroPOS")
        If Response = vbYes Then
            grdRes.Row = SaveRow
            grdRes.Col = savecol
            TillData.Res_No = Val(Mid(grdRes.TextMatrix(grdRes.Row, grdRes.Col), InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1, (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), "-") - 1) - (InStr(grdRes.TextMatrix(grdRes.Row, grdRes.Col), ":") + 1)))
            rptCheckin.Show
        End If
    End If
trap:
    On Error GoTo 0
End Sub
Private Sub grdRes_EnterCell()
     If grdRes.Col < 2 Then grdRes.Col = 2
End Sub
Private Sub grdRes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        lblRow = grdRes.Row
        lblCol = grdRes.Col
        grdRes.Select lblRow, lblCol, lblRow, lblCol
    Else
        
    End If
End Sub
Private Sub grdRes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Timer1.Enabled = True Then Exit Sub
    Static ColOld
    Static RowOld
    If grdRes.Rows > 28 Then
        MinusFactor = 3050
    Else
        MinusFactor = 3350
    End If
    If Int((x - MinusFactor) / 330) + 2 < 2 Then
        Col = 2
    Else
        If Int((x - MinusFactor) / 330) + 2 > 32 Then
            Col = 32
        Else
            Col = Int((x - MinusFactor) / 330) + 2
        End If
    End If
    If Int(y / 285) + grdRes.TopRow - 1 > grdRes.Rows - 1 Then
        Row = grdRes.Rows - 1
    Else
        If Int(y / 285) < 2 Then
            Row = grdRes.TopRow
        Else
            Row = Int(y / 285) + grdRes.TopRow - 2
        End If
    End If
    If Col <> ColOld Or Row <> RowOld Then
        grdRes.Select 2, ColOld, grdRes.Rows - 1, ColOld
        grdRes.CellBorder 0, 0, 0, 0, 0, 0, 0
        grdRes.Select RowOld, 2, RowOld, grdRes.Cols - 1
        grdRes.CellBorder 0, 0, 0, 0, 0, 0, 0
        grdRes.Cell(flexcpFontBold, RowOld, 0, RowOld, 1) = False
        grdRes.Cell(flexcpFontBold, 0, ColOld, 1, ColOld) = False
        grdRes.Cell(flexcpForeColor, RowOld, 0, RowOld, 1) = &H80000008
        grdRes.Cell(flexcpForeColor, 0, ColOld, 1, ColOld) = &H80000008
        
        RowOld = Row
        ColOld = Col
        
        grdRes.Select 2, Col, grdRes.Rows - 1, Col
        grdRes.CellBorder &HFFC0C0, 1, 0, 1, 0, 1, 0
        grdRes.Select Row, 2, Row, grdRes.Cols - 1
        grdRes.CellBorder &HFFC0C0, 0, 1, 0, 1, 0, 1
        grdRes.Cell(flexcpFontBold, Row, 0, Row, 1) = True
        grdRes.Cell(flexcpFontBold, 0, Col, 1, Col) = True
        grdRes.Cell(flexcpForeColor, RowOld, 0, RowOld, 1) = &HC0&
        grdRes.Cell(flexcpForeColor, 0, ColOld, 1, ColOld) = &HC0&
        grdRes.Col = Col
        grdRes.Row = Row
        If grdRes.TextMatrix(Row, Col) = "" Then
            If Col > 1 Then
                grdRes.ToolTipText = " Double Click for a New Reservation "
                lblType.Caption = ""
            Else
                grdRes.ToolTipText = ""
                lblType.Caption = ""
            End If
        Else
            If Col < 2 Or Row < 3 Then
                grdRes.ToolTipText = ""
            Else
                If Trim(grdRes.TextMatrix(Row, Col)) <> "" Then
                    If InStrRev(grdRes.TextMatrix(Row, Col), "|") > 1 Then
                        StringA = Mid(grdRes.TextMatrix(Row, Col), InStrRev(grdRes.TextMatrix(Row, Col), "|"))
                        stringD = Mid(grdRes.TextMatrix(Row, Col), 1, InStrRev(grdRes.TextMatrix(Row, Col), "|") - 1)
                        If InStr(StringA, "D") = 0 Then
                            SaveString = StringA
                            StringA = stringD
                            stringD = SaveString
                        End If
                        grdRes.ToolTipText = " Reservation No: " & Val(Mid(StringA, InStr(StringA, ":") + 1, (InStr(StringA, "-") - 1) - (InStr(StringA, ":") + 1))) & " > " & Mid(StringA, InStrRev(StringA, "-") + 1) & " "
                        grdRes.ToolTipText = grdRes.ToolTipText & " «»  Reservation No: " & Val(Mid(stringD, InStr(stringD, ":") + 1, (InStr(stringD, "-") - 1) - (InStr(stringD, ":") + 1))) & " < " & Mid(stringD, InStrRev(stringD, "-") + 1) & " "
                         Select Case Val(Replace(Mid(StringA, 1, InStr(StringA, "<") - 1), "|", ""))
                            Case 0: lblType.Caption = "Provisional"
                            Case 1: lblType.Caption = "Confirmed"
                            Case 2: lblType.Caption = "Checked In"
                            Case 3: lblType.Caption = "Checked Out"
                        End Select
                        Select Case Val(Replace(Mid(stringD, 1, InStr(stringD, "<") - 1), "|", ""))
                            Case 0: lblType.Caption = lblType.Caption & " «»  Provisional"
                            Case 1: lblType.Caption = lblType.Caption & " «»  Confirmed"
                            Case 2: lblType.Caption = lblType.Caption & " «»  Checked In"
                            Case 3: lblType.Caption = lblType.Caption & " «»  Checked Out"
                        End Select
                    Else
                        grdRes.ToolTipText = " Reservation No: " & Val(Mid(grdRes.TextMatrix(Row, Col), InStr(grdRes.TextMatrix(Row, Col), ":") + 1, (InStr(grdRes.TextMatrix(Row, Col), "-") - 1) - (InStr(grdRes.TextMatrix(Row, Col), ":") + 1))) & " > " & Mid(grdRes.TextMatrix(Row, Col), InStrRev(grdRes.TextMatrix(Row, Col), "-") + 1) & " "
                        Select Case Val(Replace(Mid(grdRes.TextMatrix(Row, Col), 1, InStr(grdRes.TextMatrix(Row, Col), "<") - 1), "|", ""))
                            Case 0: lblType.Caption = "Provisional"
                            Case 1: lblType.Caption = "Confirmed"
                            Case 2: lblType.Caption = "Checked In"
                            Case 3: lblType.Caption = "Checked Out"
                        End Select
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub cmdRoomRates_Click()
    Load frmRates
    frmRates.Tag = ""
    frmRates.Show vbModal
End Sub

Private Sub grdTask_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Select Case grdTask.ValueMatrix(NewRow, 3)
        Case 0
            cmdDone.Enabled = True
            cmdCanTask.Enabled = True
        Case 1
            cmdDone.Enabled = False
            cmdCanTask.Enabled = False
        Case 2
            cmdDone.Enabled = False
            cmdCanTask.Enabled = False
    End Select
    If NewRow = 0 Then
        cmdDone.Enabled = False
        cmdCanTask.Enabled = False
    End If
    If grdTask.TextMatrix(0, 1) <> "" Then
        Cmdadd.Enabled = True
    End If
End Sub
Private Sub mtView_CloseUp()
    grdRes.Rows = 3
    grdRes.RowHidden(2) = True
    grdRes.Width = 13605
    Image3.Width = 13605
    grdTask.Visible = False
    cmdTask.Caption = "Show Tasks"
    grdRes.ColWidth(0) = 2750
    grdRes.ColWidth(1) = 600
    grdRes.TextMatrix(0, 0) = "Room Description"
    grdRes.TextMatrix(0, 1) = "No."
    grdRes.ColAlignment(0) = flexAlignLeftCenter
    grdRes.ColAlignment(1) = flexAlignLeftCenter
    ActiveReadServer "select * from Rooms order by convert(int,replace(replace(replace(room_No,'a',''),'b',''),'c',''))"
    While Not rs.EOF
        grdRes.Rows = grdRes.Rows + 1
        grdRes.TextMatrix(grdRes.Rows - 1, 0) = rs.Fields("Description")
        grdRes.TextMatrix(grdRes.Rows - 1, 1) = rs.Fields("Room_No")
        rs.MoveNext
    Wend
    rs.Close
    grdRes.Cell(flexcpBackColor, 1, 0, grdRes.Rows - 1, 0) = &HFEF0EB
    grdRes.Cell(flexcpBackColor, 1, 1, grdRes.Rows - 1, 1) = &HDFFCFD
    Select Case mtView.Month
        Case 1: Lastday = 31
        Case 2
            If Int(mtView.Year / 4) = (mtView.Year / 4) Then
                Lastday = 29
            Else
                Lastday = 28
            End If
        Case 3: Lastday = 31
        Case 4: Lastday = 30
        Case 5: Lastday = 31
        Case 6: Lastday = 30
        Case 7: Lastday = 31
        Case 8: Lastday = 31
        Case 9: Lastday = 30
        Case 10: Lastday = 31
        Case 11: Lastday = 30
        Case 12: Lastday = 31
    End Select
    CurrentDate = mtView.Value
    For i = 2 To 32
        grdRes.ColWidth(i) = 330
        If i - 1 = mtView.Day Then
            Marker = i
        End If
        If i - 1 > mtView.Day Then
            grdRes.TextMatrix(2, i) = DateAdd("d", 1, CurrentDate)
            CurrentDate = DateAdd("d", 1, CurrentDate)
            grdRes.TextMatrix(0, i) = Format(CurrentDate, "D")
            Select Case Format(CurrentDate, "w")
                Case 1
                    grdRes.TextMatrix(1, i) = "Su"
                    grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
                Case 2: grdRes.TextMatrix(1, i) = "M"
                Case 3: grdRes.TextMatrix(1, i) = "T"
                Case 4: grdRes.TextMatrix(1, i) = "W"
                Case 5: grdRes.TextMatrix(1, i) = "Th"
                Case 6: grdRes.TextMatrix(1, i) = "F"
                Case 7
                    grdRes.TextMatrix(1, i) = "Sa"
                    grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
            End Select
        End If
        grdRes.ColAlignment(i) = flexAlignCenterCenter
    Next i
    If Marker = 32 Then
        grdRes.TextMatrix(0, 32) = 31
        grdRes.TextMatrix(2, 32) = mtView.Value
    End If
    For i = Marker To 2 Step -1
        If i <> 32 Then
            grdRes.TextMatrix(0, i) = Format(DateAdd("d", -1, grdRes.TextMatrix(2, i + 1)), "D")
            grdRes.TextMatrix(2, i) = DateAdd("d", -1, grdRes.TextMatrix(2, i + 1))
        End If
        Select Case Format(grdRes.TextMatrix(2, i), "w")
            Case 1
                grdRes.TextMatrix(1, i) = "Su"
                grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
            Case 2: grdRes.TextMatrix(1, i) = "M"
            Case 3: grdRes.TextMatrix(1, i) = "T"
            Case 4: grdRes.TextMatrix(1, i) = "W"
            Case 5: grdRes.TextMatrix(1, i) = "Th"
            Case 6: grdRes.TextMatrix(1, i) = "F"
            Case 7
                grdRes.TextMatrix(1, i) = "Sa"
                grdRes.Cell(flexcpBackColor, 2, i, grdRes.Rows - 1, i) = &HCCFBFB
        End Select
    Next i
    grdRes.Cell(flexcpBackColor, 2, Marker, grdRes.Rows - 1, Marker) = &HFDE1D7
    If grdRes.Rows > 28 Then
        grdRes.ColWidth(0) = 2450
    End If
    LoadRes
    Image1.Width = grdRes.ColWidth(0) + grdRes.ColWidth(1)
    Image3.Left = grdRes.ColWidth(0) + grdRes.ColWidth(1) + 90
    Image3.Width = grdRes.Width - (grdRes.ColWidth(0) + grdRes.ColWidth(1) + 30)
    cmdLeft.Left = Image3.Left + 50
    cmdRight.Left = Image3.Left + Image3.Width - 640
    cmdLeft1.Left = cmdLeft.Left + 600
    cmdRight1.Left = cmdRight.Left - 600
End Sub
Private Sub Timer1_Timer()
    Static Status
    Select Case Status
        Case 1
            Status = 0
            grdRes.Select rowStart, colStart, rowStart, colStop
            grdRes.CellBorder vbRed, 1, 1, 1, 1, 1, 1
        Case Else
            Status = 1
            grdRes.Select rowStart, colStart, rowStart, colStop
            grdRes.CellBorder 0, 0, 0, 0, 0, 0, 0
    End Select
End Sub
